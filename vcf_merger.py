import os
import re
import json
import logging
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

# Constants
CONFIG_FILE = 'vcf_config.json'
VCF_VERSION = '3.0'
MAX_PHONE_NUMBERS = 4
MAX_EMAIL_ADDRESSES = 4

def _safe_filename(name: str) -> str:
    """Create a filesystem-safe filename fragment from a contact name."""
    if not name:
        return "contact"
    # Replace invalid characters and collapse spaces
    cleaned = re.sub(r'[\\/:*?"<>|]+', '_', name)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    # Limit length to avoid path issues
    return cleaned[:80] or "contact"

def setup_logging(log_level: str = 'INFO') -> logging.Logger:
    """Setup logging configuration."""
    logger = logging.getLogger('vcf_merger')
    logger.setLevel(getattr(logging, log_level.upper()))
    
    # Remove existing handlers to avoid duplicates
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # Create file handler
    file_handler = logging.FileHandler('vcf_merger.log', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Create console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger

class VCFConfig:
    """Configuration management for VCF merger."""
    
    def __init__(self, config_file: str = CONFIG_FILE):
        self.config_file = config_file
        self.config = self.load_config()
    
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from JSON file."""
        default_config = {
            'input_files': {
                'source': 'Private Kontakte nach Aufräumen generiert aus iCloud.vcf',
                'update': 'contacts_updated.vcf'
            },
            'output_file': 'contacts_final.vcf',
            'backup_enabled': True,
            'backup_suffix': '_backup',
            'log_level': 'INFO',
            'max_phone_numbers': MAX_PHONE_NUMBERS,
            'max_email_addresses': MAX_EMAIL_ADDRESSES,
            'vcf_version': VCF_VERSION,
            'split_output': False,
            'split_output_dir': 'contacts_split'
        }
        
        file_exists = os.path.exists(self.config_file)
        if file_exists:
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    default_config.update(config)
            except Exception as e:
                print(f"Error loading config: {e}")
        else:
            # No external config present: update in-code defaults
            default_config['input_files']['source'] = 'contacts_private_v13.vcf'
            default_config['input_files']['update'] = 'icloud.vcf'
        # Remove legacy limit keys from defaults (kept for backward compatibility if provided in file)
        default_config.pop('max_phone_numbers', None)
        default_config.pop('max_email_addresses', None)
        
        return default_config
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value."""
        return self.config.get(key, default)
    
    def save_config(self) -> None:
        """Save configuration to JSON file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving config: {e}")

class VCFParser:
    """Handles parsing of VCF files."""
    
    def __init__(self, config: VCFConfig, logger: logging.Logger):
        self.config = config
        self.logger = logger
    
    def parse_name_field(self, line: str, current_contact: Dict[str, Any]) -> Tuple[Dict[str, Any], Optional[str]]:
        """Parse name fields (N, FN)."""
        try:
            if line.startswith('N:'):
                current_contact['N'] = line
                return current_contact, None
            elif line.startswith('FN:'):
                name = line.split(':', 1)[-1].strip()
                current_contact['FN'] = name
                return current_contact, name
        except Exception as e:
            self.logger.error(f"Error parsing name field '{line}': {e}")
        return current_contact, None
    
    def parse_birthday_field(self, raw_value: str) -> str:
        """Parse birthday field and normalize to YYYY-MM-DD. If year is missing, use 1900.

        Supports formats like:
        - YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
        - DD-MM-YYYY, DD/MM/YYYY, DD.MM.YYYY
        - DD-MM, DD/MM, DD.MM (assumed European order → 1900-MM-DD)
        - YYYYMMDD, DDMMYYYY
        """
        try:
            value = raw_value.split(':', 1)[-1].strip() if ':' in raw_value else raw_value.strip()
            if not value:
                return value

            # Helper to validate date parts
            def valid_ymd(y: int, m: int, d: int) -> bool:
                return 1 <= m <= 12 and 1 <= d <= 31 and 1800 <= y <= 2200

            # 1) YYYY[-./]MM[-./]DD
            m = re.match(r'^(\d{4})[\-\./](\d{1,2})[\-\./](\d{1,2})$', value)
            if m:
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if valid_ymd(y, mo, d):
                    return f"{y:04d}-{mo:02d}-{d:02d}"

            # 2) DD[-./]MM[-./]YYYY (European)
            m = re.match(r'^(\d{1,2})[\-\./](\d{1,2})[\-\./](\d{4})$', value)
            if m:
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if valid_ymd(y, mo, d):
                    return f"{y:04d}-{mo:02d}-{d:02d}"

            # 3) DD[-./]MM (no year → 1900)
            m = re.match(r'^(\d{1,2})[\-\./](\d{1,2})$', value)
            if m:
                d, mo = int(m.group(1)), int(m.group(2))
                if 1 <= mo <= 12 and 1 <= d <= 31:
                    return f"1900-{mo:02d}-{d:02d}"

            # 4) YYYYMMDD
            m = re.match(r'^(\d{4})(\d{2})(\d{2})$', value)
            if m:
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if valid_ymd(y, mo, d):
                    return f"{y:04d}-{mo:02d}-{d:02d}"

            # 5) DDMMYYYY (European compact)
            m = re.match(r'^(\d{2})(\d{2})(\d{4})$', value)
            if m:
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if valid_ymd(y, mo, d):
                    return f"{y:04d}-{mo:02d}-{d:02d}"

            # If already in correct form, return as-is
            if re.match(r'^\d{4}-\d{2}-\d{2}$', value):
                return value

            # Fallback: return original value
            return value
        except Exception as e:
            self.logger.error(f"Error parsing birthday field '{raw_value}': {e}")
            return raw_value

    def _is_valid_phone_number(self, phone: str, current_contact: Dict[str, Any]) -> bool:
        """Validate phone number and check for duplicates."""
        # Get configuration values
        min_digits = self.config.get('phone_validation', {}).get('min_digits', 7)
        check_duplicates = self.config.get('phone_validation', {}).get('check_duplicates', True)
        
        # Remove all non-digit characters for length check
        digits_only = re.sub(r'[^\d]', '', phone)
        
        # Must have at least minimum digits
        if len(digits_only) < min_digits:
            self.logger.debug(f"Phone number too short: {phone} (only {len(digits_only)} digits, need {min_digits})")
            return False
        
        # Must not be empty or just zeros
        if not phone or phone == '0' or phone == '':
            return False
        
        # Additional validation: must contain at least one non-zero digit
        if digits_only == '0' * len(digits_only):
            self.logger.debug(f"Phone number contains only zeros: {phone}")
            return False
        
        # Check for duplicates if enabled
        if check_duplicates:
            normalized_phone = re.sub(r'[^\d]', '', phone)
            existing_phones = current_contact.get('TEL', [])
            
            for existing_tel in existing_phones:
                if ':' in existing_tel:
                    existing_phone = existing_tel.split(':', 1)[-1]
                    existing_normalized = re.sub(r'[^\d]', '', existing_phone)
                    if normalized_phone == existing_normalized:
                        self.logger.debug(f"Duplicate phone number found: {phone}")
                        return False
        
        return True

    def parse_phone_field(self, line: str, current_contact: Dict[str, Any], current_name: Optional[str]) -> None:
        """Parse phone number fields with comprehensive format support."""
        try:
            if 'TEL' not in current_contact or not isinstance(current_contact['TEL'], list):
                current_contact['TEL'] = []
            
            # Handle iCloud/typed format: TEL;type=CELL;type=VOICE;type=pref:+49 ...
            if line.startswith('TEL;') and 'type=' in line:
                phone_match = re.search(r':\s*([+\d\s\(\)\-\.]+)\s*$', line)
                if phone_match:
                    phone = phone_match.group(1).strip()
                    if self._is_valid_phone_number(phone, current_contact):
                        # Preserve original TYPE parameters for Outlook mapping
                        before_colon = line.split(':', 1)[0]
                        param_segment = before_colon[len('TEL'):]
                        current_contact['TEL'].append(f"TEL{param_segment}:{phone}")
                        self.logger.debug(f"Found typed phone: {line} -> Preserved types for {current_name}")
                        return
                    else:
                        self.logger.debug(f"Phone validation failed for: {phone} from line: {line}")
            
            # Handle standard format: TEL:+4917642249602
            elif line.startswith('TEL:'):
                phone = line.split(':', 1)[-1].strip()
                if self._is_valid_phone_number(phone, current_contact):
                    current_contact['TEL'].append(f"TEL:{phone}")
                    self.logger.debug(f"Found standard phone: {line} -> Extracted: TEL:{phone} for {current_name}")
                    return
            
            # Handle item format: item1.TEL;type=CELL;type=VOICE;type=pref:+49 ...
            elif line.startswith('item') and 'TEL' in line:
                phone_match = re.search(r':\s*([+\d\s\(\)\-\.]+)\s*$', line)
                if phone_match:
                    phone = phone_match.group(1).strip()
                    if self._is_valid_phone_number(phone, current_contact):
                        # Convert itemX.TEL;... to TEL;...
                        after_dot = line.split('.', 1)[-1]
                        before_colon = after_dot.split(':', 1)[0]
                        if before_colon.upper().startswith('TEL'):
                            param_segment = before_colon[len('TEL'):]
                        else:
                            param_segment = ''
                        current_contact['TEL'].append(f"TEL{param_segment}:{phone}")
                        self.logger.debug(f"Found item phone: {line} -> Preserved as TEL{param_segment}:{phone} for {current_name}")
                        return
            
            # Handle TEL with TYPE= parameters
            elif line.startswith('TEL;') and 'TYPE=' in line.upper():
                phone_match = re.search(r':\s*([+\d\s\(\)\-\.]+)\s*$', line)
                if phone_match:
                    phone = phone_match.group(1).strip()
                    if self._is_valid_phone_number(phone, current_contact):
                        before_colon = line.split(':', 1)[0]
                        param_segment = before_colon[len('TEL'):]
                        current_contact['TEL'].append(f"TEL{param_segment}:{phone}")
                        self.logger.debug(f"Found TYPE phone: {line} -> Preserved types for {current_name}")
                        return
            
            # Fallback: try to extract any phone-like pattern - improved regex
            else:
                # More comprehensive phone number pattern - look for longer sequences
                phone_match = re.search(r'([+\d\s\(\)\-\.]{10,})', line)
                if phone_match:
                    phone = phone_match.group(1).strip()
                    if self._is_valid_phone_number(phone, current_contact):
                        current_contact['TEL'].append(f"TEL:{phone}")
                        self.logger.debug(f"Found fallback phone: {line} -> Extracted: TEL:{phone} for {current_name}")
                        return
            
            # If no valid phone found, log it with more detail
            if current_name:
                self.logger.warning(f"No valid phone number extracted from: {line} for {current_name}")
            else:
                self.logger.warning(f"No valid phone number extracted from: {line}")
            
        except Exception as e:
            self.logger.error(f"Error parsing phone field '{line}' for {current_name}: {e}")

    def parse_email_field(self, line: str, value: str, current_contact: Dict[str, Any], current_name: Optional[str]) -> None:
        """Parse email fields."""
        if 'EMAIL' not in current_contact or not isinstance(current_contact['EMAIL'], list):
            current_contact['EMAIL'] = []
        
        email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', value)
        if email_match:
            # Store emails in standard form for vCard writing
            current_contact['EMAIL'].append(f"EMAIL:{email_match.group(1)}")
            self.logger.debug(f"Found email: {line} -> Extracted: EMAIL:{email_match.group(1)} for {current_name}")
        else:
            normalized_value = value.rstrip(';').strip()
            current_contact['EMAIL'].append(f"EMAIL:{normalized_value}")
            self.logger.debug(f"Found email: {line} -> No valid email found, using: EMAIL:{normalized_value} for {current_name}")
    
    def extract_emails_from_notes(self, current_contact: Dict[str, Any], current_name: Optional[str]) -> None:
        """
        CRITICAL FUNCTION: Extract email addresses from NOTE fields.
        
        This method is essential for processing iCloud VCF files where emails are stored in NOTE fields
        instead of EMAIL fields. It searches for various email patterns and extracts them.
        
        WARNING: DO NOT REMOVE OR SIMPLIFY THIS METHOD - it's critical for proper email extraction!
        
        Patterns searched:
        - 'E-mail Address:' (e.g., 'E-mail Address: Angelika.GRIX@3ds.com')
        - 'E-mail 2 Address:' (e.g., 'E-mail 2 Address: doris.helfinger@gmx.de')
        - 'E-mail Display Name:' (e.g., 'E-mail Display Name: GRIX Angelika')
        
        If this method is removed or broken, contacts like Angelika Grix will lose their emails!
        """
        if 'NOTE' not in current_contact or not isinstance(current_contact['NOTE'], list):
            self.logger.debug(f"No NOTE field found for {current_name}, skipping email extraction")
            return
        
        if 'EMAIL' not in current_contact or not isinstance(current_contact['EMAIL'], list):
            current_contact['EMAIL'] = []
        
        emails_extracted = 0
        remaining_notes: List[str] = []
        for note in current_contact['NOTE']:
            # CRITICAL: Look for various email patterns in NOTE fields
            # These patterns are specific to iCloud VCF format
            if any(pattern in note for pattern in ['E-mail Address:', 'E-mail 2 Address:', 'E-mail Display Name:']):
                # Extract email using regex - this is the core functionality
                email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', note)
                if email_match:
                    email = email_match.group(1)
                    self.logger.info(f"Found email in NOTE for {current_name}: {email}")
                    
                    # Check if email already exists to avoid duplicates
                    existing_emails = [e.split(':', 1)[-1] for e in current_contact['EMAIL'] if ':' in e]
                    if email not in existing_emails:
                        # Store in standard vCard form (repeatable EMAIL field) without limiting count
                        email_field = f"EMAIL:{email}"
                        current_contact['EMAIL'].append(email_field)
                        emails_extracted += 1
                        self.logger.info(f"SUCCESS: Extracted email from NOTE: {email} -> {email_field} for {current_name}")
                    else:
                        self.logger.debug(f"Email {email} already exists for {current_name}, skipping duplicate")
                    # We successfully extracted or recognized a duplicate → drop this NOTE line
                    continue
                else:
                    self.logger.warning(f"Email pattern found in NOTE but no valid email extracted: {note}")
                    remaining_notes.append(note)
            else:
                self.logger.debug(f"No email pattern found in NOTE: {note}")
                remaining_notes.append(note)
        
        # Keep only non-redundant NOTE lines
        current_contact['NOTE'] = remaining_notes

        if emails_extracted > 0:
            self.logger.info(f"Total emails extracted from NOTES for {current_name}: {emails_extracted}")
        else:
            self.logger.debug(f"No emails extracted from NOTES for {current_name}")

    def extract_phones_from_notes(self, current_contact: Dict[str, Any], current_name: Optional[str]) -> None:
        """
        Extract phone numbers from NOTE fields and add them as TEL entries.

        Looks for common Outlook/iCloud note labels like:
        - 'Business Phone:', 'Home Phone:', 'Mobile Phone:', 'Other Phone:'
        And also scans for generic phone-like patterns within NOTE lines.
        Respects MAX_PHONE_NUMBERS and avoids duplicates using existing validation.
        """
        if 'NOTE' not in current_contact or not isinstance(current_contact['NOTE'], list):
            self.logger.debug(f"No NOTE field found for {current_name}, skipping phone extraction")
            return

        if 'TEL' not in current_contact or not isinstance(current_contact['TEL'], list):
            current_contact['TEL'] = []

        extracted = 0
        # Regex to find phone-like sequences; we'll still validate via _is_valid_phone_number
        phone_pattern = re.compile(r"(\+?\d[\d\s().\-]{6,}\d)")

        remaining_notes: List[str] = []
        for note in current_contact['NOTE']:
            # Prefer text after a known phone label if present
            label = None
            for lbl in ['Business Phone:', 'Home Phone:', 'Mobile Phone:', 'Other Phone:', 'Phone:']:
                if lbl in note:
                    label = lbl
                    break
            if label:
                labeled_part = note.split(':', 1)[-1]
                candidates = phone_pattern.findall(labeled_part)
            else:
                candidates = phone_pattern.findall(note)

            added_from_this_note = False
            for cand in candidates:
                phone = cand.strip()
                if self._is_valid_phone_number(phone, current_contact):
                    # Avoid duplicates
                    existing_nums = []
                    for existing_tel in current_contact.get('TEL', []):
                        if ':' in existing_tel:
                            existing_nums.append(re.sub(r'[^\d]', '', existing_tel.split(':', 1)[-1]))
                    normalized_new = re.sub(r'[^\d]', '', phone)
                    if normalized_new in existing_nums:
                        continue
                    # Add TEL with TYPE based on label if known
                    type_param = ''
                    if label:
                        if 'Business' in label:
                            type_param = ';TYPE=WORK;TYPE=VOICE'
                        elif 'Home' in label:
                            type_param = ';TYPE=HOME;TYPE=VOICE'
                        elif 'Mobile' in label:
                            type_param = ';TYPE=CELL;TYPE=VOICE'
                        elif 'Other' in label:
                            type_param = ';TYPE=VOICE'
                    current_contact['TEL'].append(f"TEL{type_param}:{phone}")
                    extracted += 1
                    added_from_this_note = True
            # Keep the note only if nothing was extracted from it
            if not added_from_this_note:
                remaining_notes.append(note)
        # Replace NOTE list with remaining (non-redundant) entries
        current_contact['NOTE'] = remaining_notes
        if extracted:
            self.logger.info(f"Extracted {extracted} phone number(s) from NOTES for {current_name}")
        else:
            self.logger.debug(f"No phone numbers extracted from NOTES for {current_name}")

    def cleanup_notes(self, current_contact: Dict[str, Any], current_name: Optional[str]) -> None:
        """Remove NOTE lines that duplicate data already promoted to structured fields.

        Drops NOTE lines like:
        - Job Title: ... (if TITLE present)
        - Business/Home/Mobile/Other Phone: ... (if TEL exists)
        - Business Street/City/Postal Code/Country: ... (if ADR exists)
        - E-mail Address/Type/Display Name: ... (if EMAIL exists)
        Keeps all other NOTE lines.
        """
        notes = current_contact.get('NOTE')
        if not isinstance(notes, list):
            return

        has_tel = bool(current_contact.get('TEL'))
        has_adr = bool(current_contact.get('ADR'))
        has_email = bool(current_contact.get('EMAIL'))
        has_title = bool(current_contact.get('TITLE'))

        drop_prefixes = []
        if has_title:
            drop_prefixes.append('NOTE:Job Title:')
        if has_tel:
            drop_prefixes += [
                'NOTE:Business Phone:', 'NOTE:Home Phone:', 'NOTE:Mobile Phone:', 'NOTE:Other Phone:', 'NOTE:Phone:'
            ]
        if has_adr:
            drop_prefixes += [
                'NOTE:Business Street:', 'NOTE:Business City:', 'NOTE:Business Postal Code:', 'NOTE:Business Country/Region:'
            ]
        if has_email:
            drop_prefixes += [
                'NOTE:E-mail Address:', 'NOTE:E-mail Type:', 'NOTE:E-mail Display Name:'
            ]
        # Always drop these regardless of extracted fields
        drop_prefixes += [
            'NOTE:Priority:', 'NOTE:Sensitivity:'
        ]

        remaining: List[str] = []
        for note in notes:
            if any(note.startswith(pfx) for pfx in drop_prefixes):
                continue
            remaining.append(note)
        current_contact['NOTE'] = remaining
        self.logger.debug(f"Cleanup NOTES for {current_name}: kept {len(remaining)} note lines")

    def parse_address_field(self, key: str, value: str, current_contact: Dict[str, Any], current_name: Optional[str]) -> None:
        """Parse address fields, preserving TYPE parameters when present."""
        if 'ADR' not in current_contact or not isinstance(current_contact['ADR'], list):
            current_contact['ADR'] = []
        
        adr_parts = value.split(';')
        while len(adr_parts) < 7:
            adr_parts.append('')
        # Preserve TYPE from key if present, including itemX.ADR;...
        param_segment = ''
        k = key
        if k.lower().startswith('item') and '.adr' in k.lower():
            k = k.split('.', 1)[-1]  # drop itemX.
        if k.upper().startswith('ADR;'):
            param_segment = k[3:]  # keep ;TYPE=...
        adr_entry = f"ADR{param_segment}:{';'.join(adr_parts[:7])}"
        current_contact['ADR'].append(adr_entry)
        self.logger.debug(f"Found address: {adr_entry} for {current_name}")

    def parse_vcard_line(self, line: str, current_name: Optional[str], current_contact: Dict[str, Any]) -> Tuple[Dict[str, Any], Optional[str]]:
        """Parse a single vCard line."""
        line = line.strip()
        self.logger.debug(f"Processing line: {line}")
        
        if line == 'BEGIN:VCARD':
            return {}, None
        elif line.startswith('END:VCARD'):
            if current_name and current_contact:
                # Extract emails from NOTE fields before completing vCard
                self.extract_emails_from_notes(current_contact, current_name)
                # Also extract phone numbers from NOTE fields
                self.extract_phones_from_notes(current_contact, current_name)
                # Remove redundant NOTE lines
                self.cleanup_notes(current_contact, current_name)
                self.logger.debug(f"vCard for {current_name} completed, fields before processing: {current_contact}")
            return current_contact, current_name
        
        # Parse name fields
        current_contact, name = self.parse_name_field(line, current_contact)
        if name:
            return current_contact, name
        
        # Parse other fields
        if ':' in line and not line.startswith('END:'):
            key, value = line.split(':', 1)
            normalized_value = value.rstrip(';').strip()
            
            if key.startswith('BDAY'):
                if 'BDAY' not in current_contact:
                    normalized_value = self.parse_birthday_field(normalized_value)
                    current_contact['BDAY'] = normalized_value
                    self.logger.debug(f"Read BDAY field: {normalized_value} for {current_name}")
            elif key.startswith('TEL'):
                self.parse_phone_field(line, current_contact, current_name)
            elif re.search(r'^(?:item\d+\.)?EMAIL', key, re.IGNORECASE):
                self.parse_email_field(line, value, current_contact, current_name)
            elif re.search(r'^(?:item\d+\.)?ADR', key, re.IGNORECASE):
                self.parse_address_field(key, value, current_contact, current_name)
            elif key == 'TITLE':
                # Store proper TITLE field instead of duplicating into NOTE
                current_contact['TITLE'] = normalized_value
                self.logger.debug(f"Found job title (TITLE): {normalized_value} for {current_name}")
            elif key in ['NOTE', 'ORG']:
                if key in current_contact and not isinstance(current_contact[key], list):
                    current_contact[key] = [current_contact[key]]
                current_contact[key] = current_contact.get(key, []) + [line]
            else:
                current_contact[key] = normalized_value
                if current_name:
                    self.logger.debug(f"Read field {key}: {normalized_value} for {current_name}")
        
        return current_contact, current_name

class VCFProcessor:
    """Handles VCF file processing and contact merging."""
    
    def __init__(self, config: VCFConfig, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.parser = VCFParser(config, logger)
    
    def read_vcf(self, vcf_file: str) -> Dict[str, Dict[str, Any]]:
        """Read VCF file and return contacts dictionary."""
        self.logger.info(f"Opening file: {vcf_file}")
        contacts = {}
        current_contact = {}
        current_name = None
        line_count = 0
        
        try:
            with open(vcf_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line_count += 1
                    if line_count % 2000 == 0:
                        self.logger.info(f"Processed {line_count} lines...")
                    
                    line = line.strip()
                    if not line:
                        continue
                    
                    current_contact, name = self.parser.parse_vcard_line(line, current_name, current_contact)
                    
                    if name and name != current_name:
                        current_name = name
                    
                    if line.startswith('END:VCARD') and current_name and current_contact:
                        # CRITICAL: Extract emails from NOTE fields BEFORE saving the contact
                        # This ensures that both source and update files have their emails extracted
                        # WARNING: DO NOT REMOVE THIS - it's essential for contacts like Angelika Grix!
                        self.parser.extract_emails_from_notes(current_contact, current_name)
                        # Also extract phone numbers from NOTE fields
                        self.parser.extract_phones_from_notes(current_contact, current_name)
                        # Remove redundant NOTE lines
                        self.parser.cleanup_notes(current_contact, current_name)
                        self.logger.debug(f"Contact completed for {current_name}, emails after extraction: {current_contact.get('EMAIL', [])}")
                        contacts[current_name] = current_contact.copy()
                        current_contact = {}
                        current_name = None
                        
        except UnicodeDecodeError:
            # Try with Latin-1 encoding
            with open(vcf_file, 'r', encoding='latin-1') as f:
                for line in f:
                    line_count += 1
                    if line_count % 2000 == 0:
                        self.logger.info(f"Processed {line_count} lines...")
                    
                    line = line.strip()
                    if not line:
                        continue
                    
                    current_contact, name = self.parser.parse_vcard_line(line, current_name, current_contact)
                    
                    if name and name != current_name:
                        current_name = name
                    
                    if line.startswith('END:VCARD') and current_name and current_contact:
                        # CRITICAL: Extract emails from NOTE fields BEFORE saving the contact
                        # This ensures that both source and update files have their emails extracted
                        # WARNING: DO NOT REMOVE THIS - it's essential for contacts like Angelika Grix!
                        self.parser.extract_emails_from_notes(current_contact, current_name)
                        # Also extract phone numbers from NOTE fields
                        self.parser.extract_phones_from_notes(current_contact, current_name)
                        # Remove redundant NOTE lines
                        self.parser.cleanup_notes(current_contact, current_name)
                        self.logger.debug(f"Contact completed for {current_name}, emails after extraction: {current_contact.get('EMAIL', [])}")
                        contacts[current_name] = current_contact.copy()
                        current_contact = {}
                        current_name = None
        except Exception as e:
            self.logger.error(f"Error reading {vcf_file}: {e}")
            raise
        
        self.logger.info(f"Read {len(contacts)} contacts from {vcf_file}")
        return contacts

    def _auto_resolve_conflict(self, field: str, source_value: Any, update_value: Any, source_normalized: Any, update_normalized: Any) -> str:
        """
        Automatically resolves conflicts based on field type and configuration.
        Returns 'update', 'source', or 'merge'.
        """
        # Get configuration preferences
        prefer_update_for = self.config.get('conflict_resolution', {}).get('prefer_update_for', ['EMAIL', 'TEL', 'ADR', 'ORG', 'NOTE'])
        prefer_source_for = self.config.get('conflict_resolution', {}).get('prefer_source_for', ['N', 'FN', 'BDAY'])
        
        # BDAY fields: Prefer source (usually more accurate, avoid 1900-01-01)
        if field == 'BDAY':
            if '1900-01-01' in str(update_normalized) or '1900-01-01' in str(update_value):
                return 'source'  # Keep source if update has default date
            elif '1900-01-01' in str(source_normalized) or '1900-01-01' in str(source_value):
                return 'update'  # Use update if source has default date
            else:
                return 'source'  # Default: prefer source for BDAY
        
        # Merge list fields where combining makes sense
        if isinstance(source_value, list) and isinstance(update_value, list):
            if field in ['EMAIL', 'TEL', 'ADR', 'NOTE']:
                return 'merge'

        # Use configuration preferences
        elif field in prefer_update_for:
            return 'update'
        elif field in prefer_source_for:
            return 'source'

        # For list fields, merge them
        elif isinstance(source_value, list) and isinstance(update_value, list):
            return 'merge'
        
        # Default: prefer update (more current information)
        return 'update'

    def normalize_value(self, val: Any) -> Any:
        """Normalize value for comparison."""
        if isinstance(val, str):
            parts = val.split(':', 1)[-1].split(';')
            return ' '.join([p.strip() for p in parts if p.strip() and not p.startswith('type=')])
        elif isinstance(val, list):
            return [self.normalize_value(v) for v in val]
        return val

    def merge_contacts(self, source_data: Dict[str, Any], update_data: Dict[str, Any]) -> Dict[str, Any]:
        """Merge contact information with automatic conflict resolution."""
        merged = source_data.copy()
        
        for field, update_value in update_data.items():
            source_value = merged.get(field)
            self.logger.debug(f"Comparing field {field}: Source: {source_value}, Update: {update_value}")
            
            if source_value and update_value:
                source_normalized = self.normalize_value(source_value)
                update_normalized = self.normalize_value(update_value)
                self.logger.debug(f"Normalized: Source: {source_normalized}, Update: {update_normalized}")
                
                # Use automatic conflict resolution
                resolution = self._auto_resolve_conflict(field, source_value, update_value, source_normalized, update_normalized)
                
                if resolution == 'update':
                    merged[field] = update_value
                    self.logger.info(f"AUTO-RESOLVED: Update value chosen for {field}: {update_normalized}")
                elif resolution == 'source':
                    self.logger.info(f"AUTO-RESOLVED: Source value kept for {field}: {source_normalized}")
                elif resolution == 'merge':
                    merged[field] = list(set(source_value + update_value))
                    self.logger.info(f"AUTO-RESOLVED: Values merged for {field}")
            elif update_value:
                # CRITICAL: For EMAIL fields, merge instead of overwrite
                # This prevents losing emails extracted from NOTE fields
                if field == 'EMAIL' and isinstance(update_value, list):
                    if 'EMAIL' not in merged:
                        merged[field] = []
                    # Add new emails without overwriting existing ones
                    for email in update_value:
                        if email not in merged[field]:
                            merged[field].append(email)
                    self.logger.debug(f"Merged EMAIL field for {merged.get('FN', 'Unknown')}: {len(merged[field])} emails")
                else:
                    merged[field] = update_value
                    self.logger.debug(f"Added {field} with: {update_value}")
        
        # CRITICAL: Extract emails from NOTE fields after merging
        # This ensures that emails stored in NOTE fields (like in iCloud VCF files)
        # are properly extracted and added to EMAIL fields
        # WARNING: DO NOT REMOVE THIS - it's essential for contacts like Angelika Grix!
        if 'NOTE' in merged and isinstance(merged['NOTE'], list):
            self.logger.info(f"Extracting emails/phones from NOTES for merged contact: {merged.get('FN', 'Unknown')}")
            self.parser.extract_emails_from_notes(merged, merged.get('FN', 'Unknown'))
            self.parser.extract_phones_from_notes(merged, merged.get('FN', 'Unknown'))
        else:
            self.logger.debug(f"No NOTE field found for email extraction in merged contact: {merged.get('FN', 'Unknown')}")
        
        return merged

    def get_contact_key(self, contact: Dict[str, Any]) -> str:
        """Generate a normalized key for a contact."""
        name = contact.get('FN', '')
        if not name:
            name = contact.get('N', '').split(';')[0] if contact.get('N') else ''
        
        # Normalize name by sorting parts
        name_parts = re.split(r'[,\s]+', name.strip())
        name_parts = [part.strip() for part in name_parts if part.strip()]
        return ' '.join(sorted(name_parts)).lower()

    def calculate_completeness_score(self, contact: Dict[str, Any]) -> int:
        """Calculate completeness score for a contact."""
        score = 0
        
        # Basic fields
        if contact.get('FN'): score += 1
        if contact.get('N'): score += 1
        if contact.get('ORG'): score += 1
        if contact.get('BDAY') and contact.get('BDAY') != '1900-01-01': score += 1
        
        # Phone numbers
        tel_count = len([tel for tel in contact.get('TEL', []) if ':' in tel and tel.split(':', 1)[-1].strip()])
        score += min(tel_count, 3)
        
        # Email addresses
        email_count = len([email for email in contact.get('EMAIL', []) if ':' in email and email.split(':', 1)[-1].strip()])
        score += min(email_count, 3)
        
        # Address
        if contact.get('ADR'): score += 1
        
        # Notes
        if contact.get('NOTE'): score += 1
        
        return score

    def remove_duplicates(self, contacts: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """Remove duplicate contacts, keeping the most complete entry."""
        self.logger.info("Removing duplicate contacts...")
        
        contact_groups = {}
        for name, contact in contacts.items():
            key = self.get_contact_key(contact)
            if key not in contact_groups:
                contact_groups[key] = []
            contact_groups[key].append((name, contact))
        
        unique_contacts = {}
        duplicates_removed = 0
        
        for key, group in contact_groups.items():
            if len(group) == 1:
                unique_contacts[group[0][0]] = group[0][1]
            else:
                # Merge duplicates instead of discarding data
                # Start with the most complete contact to preserve richer data
                best_name, best_data = max(group, key=lambda x: self.calculate_completeness_score(x[1]))
                merged_data = best_data
                for name, contact in group:
                    if contact is merged_data:
                        continue
                    merged_data = self.merge_contacts(merged_data, contact)
                # Use a stable name: prefer FN if available
                final_name = merged_data.get('FN') or best_name
                unique_contacts[final_name] = merged_data
                duplicates_removed += len(group) - 1
                self.logger.info(f"Merged {len(group)} duplicates for '{key}', kept merged contact: {final_name}")
        
        self.logger.info(f"Removed {duplicates_removed} duplicate contacts")
        return unique_contacts

    def write_vcf(self, contacts: Dict[str, Dict[str, Any]], output_vcf: str) -> None:
        """Write merged contacts to VCF file."""
        self.logger.info(f"Writing {len(contacts)} contacts to {output_vcf}")
        
        try:
            split_enabled = self.config.get('split_output', False)
            split_dir = self.config.get('split_output_dir', 'contacts_split')
            if split_enabled:
                os.makedirs(split_dir, exist_ok=True)
            with open(output_vcf, 'w', encoding='utf-8') as f:
                if contacts:
                    for i, (full_name, data) in enumerate(contacts.items(), 1):
                        lines: List[str] = []
                        lines.append("BEGIN:VCARD\n")
                        lines.append(f"VERSION:{self.config.get('vcf_version', VCF_VERSION)}\n")
                        lines.append(f"N:{data.get('N', ';;;')}\n")
                        lines.append(f"FN:{data.get('FN', full_name)}\n")
                        lines.append(f"ORG:{data.get('ORG', [''])[-1].split(':', 1)[-1] if data.get('ORG') else ''}\n")
                        if data.get('TITLE'):
                            lines.append(f"TITLE:{data.get('TITLE')}\n")
                        lines.append(f"BDAY:{data.get('BDAY', '1900-01-01')}\n")
                        
                        adr = data.get('ADR', ['ADR:;;;;;;;'])[0] if data.get('ADR') else 'ADR:;;;;;;;'
                        lines.append(f"{adr}\n")
                        
                        # Preserve original TEL entries (including TYPE params) and de-duplicate by number
                        tel_values_raw = [tel for tel in data.get('TEL', []) if ':' in tel]
                        seen_nums = set()
                        tel_entries: List[Tuple[int, int, str]] = []  # (priority, original_index, line)
                        def tel_priority(tel_line: str) -> int:
                            # Lower is better: prioritize CELL, then WORK, then HOME, then others
                            before_colon = tel_line.split(':', 1)[0].upper()
                            if 'CELL' in before_colon:
                                return 0
                            if 'WORK' in before_colon or 'BUSINESS' in before_colon:
                                return 1
                            if 'HOME' in before_colon:
                                return 2
                            if 'FAX' in before_colon:
                                return 3
                            return 4
                        for idx, tel in enumerate(tel_values_raw):
                            num = tel.split(':', 1)[-1].strip()
                            if not num:
                                continue
                            normalized = re.sub(r'[^\d]', '', num)
                            if normalized in seen_nums:
                                continue
                            seen_nums.add(normalized)
                            line = tel if tel.upper().startswith('TEL') else f"TEL:{num}"
                            tel_entries.append((tel_priority(line), idx, line))
                        # Sort by priority, keeping stable order within same priority, then write all
                        tel_entries.sort(key=lambda x: (x[0], x[1]))
                        for _, __, tel_line in tel_entries:
                            lines.append(f"{tel_line}\n")
                        
                        # Collect up to max email addresses and write as standard EMAIL entries
                        email_values_raw = [email for email in data.get('EMAIL', []) if ':' in email]
                        email_addresses = []
                        for em in email_values_raw:
                            addr = em.split(':', 1)[-1].strip()
                            if addr:
                                email_addresses.append(addr)
                        # Preserve order and de-duplicate while respecting max limit
                        seen_emails = set()
                        unique_emails = []
                        for addr in email_addresses:
                            if addr not in seen_emails:
                                seen_emails.add(addr)
                                unique_emails.append(addr)
                        for idx, addr in enumerate(unique_emails, start=1):
                            lines.append(f"EMAIL:{addr}\n")
                            # DEBUG: Log email writing for Angelika Grix
                            if 'Angelika' in full_name or 'Grix' in full_name:
                                self.logger.debug(f"DEBUG: Writing EMAIL:{addr} for {full_name}")
                        
                        # Combine all remaining notes into a single NOTE field so
                        # Outlook/iCloud display them together. Use \n escapes within the value.
                        note_values = data.get('NOTE', [])
                        if note_values:
                            contents: List[str] = []
                            for note in note_values:
                                # Strip leading "NOTE:" if present to avoid double prefixing
                                if note.upper().startswith('NOTE:'):
                                    contents.append(note.split(':', 1)[-1])
                                else:
                                    contents.append(note)
                            joined = "\\n".join([c.strip() for c in contents if c.strip()])
                            if joined:
                                # Escape backslashes to be vCard-safe
                                joined_escaped = joined.replace('\\', '\\\\')
                                lines.append(f"NOTE:{joined_escaped}\n")
                        lines.append("END:VCARD\n")

                        # Write combined file block
                        f.write(''.join(lines))

                        # Optionally write per-contact .vcf file for Outlook import
                        if split_enabled:
                            base_name = data.get('FN') or full_name or data.get('N', 'contact')
                            fname = _safe_filename(base_name)
                            path = os.path.join(split_dir, f"{fname}.vcf")
                            # Ensure uniqueness if file exists
                            if os.path.exists(path):
                                suffix = 2
                                while os.path.exists(os.path.join(split_dir, f"{fname}_{suffix}.vcf")):
                                    suffix += 1
                                path = os.path.join(split_dir, f"{fname}_{suffix}.vcf")
                            with open(path, 'w', encoding='utf-8') as single:
                                single.write(''.join(lines))
                        
                        # Progress indicator
                        if i % 50 == 0:
                            self.logger.info(f"Written {i}/{len(contacts)} contacts...")
                        
                        self.logger.debug(f"Written contact: {full_name}")
                else:
                    self.logger.warning("No contacts to write.")
            
            self.logger.info(f"Updated VCF file created: {output_vcf}")
            
        except Exception as e:
            self.logger.error(f"Error writing output file: {e}")
            raise

class VCFMerger:
    """Main class for VCF merging operations."""
    
    def __init__(self, config_file: str = CONFIG_FILE):
        self.config = VCFConfig(config_file)
        self.logger = setup_logging(self.config.get('log_level', 'INFO'))
        self.processor = VCFProcessor(self.config, self.logger)
    
    def create_backup(self, file_path: str) -> None:
        """Create backup of existing file."""
        if not self.config.get('backup_enabled', True):
            return
        
        if os.path.exists(file_path):
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_suffix = self.config.get('backup_suffix', '_backup')
            backup_path = f"{file_path}{backup_suffix}_{timestamp}"
            
            try:
                shutil.copy2(file_path, backup_path)
                self.logger.info(f"Backup created: {backup_path}")
            except Exception as e:
                self.logger.error(f"Error creating backup: {e}")
    
    def update_vcf_with_vcf(self, remove_duplicates_flag: bool = True) -> str:
        """Update VCF file with another VCF file."""
        try:
            # Get file paths
            source_file = self.config.get('input_files', {}).get('source')
            update_file = self.config.get('input_files', {}).get('update')
            output_file = self.config.get('output_file')
            
            if not source_file:
                raise ValueError("Source file must be specified in config")
            
            # Create backup of output file if it exists
            self.create_backup(output_file)
            
            # Read source VCF
            self.logger.info("Reading source VCF file...")
            source_contacts = self.processor.read_vcf(source_file)
            
            # Check if update file is provided
            if update_file:
                # Read update VCF
                self.logger.info("Reading update VCF file...")
                update_contacts = self.processor.read_vcf(update_file)
                
                self.logger.info(f"Contacts from base (iCloud): {len(source_contacts)}")
                self.logger.info(f"Contacts from update: {len(update_contacts)}")
                
                # Start with source contacts
                merged_contacts = source_contacts.copy()
                self.logger.info(f"Initial merged contacts: {len(merged_contacts)}")
                
                # Process each update contact
                update_count = len(update_contacts)
                for i, (update_name, update_data) in enumerate(update_contacts.items(), 1):
                    if i % 10 == 0:
                        self.logger.info(f"Processed {i}/{update_count} contacts...")
                    
                    if update_name in merged_contacts:
                        # Merge existing contact
                        source_data = merged_contacts[update_name]
                        merged_data = self.processor.merge_contacts(source_data, update_data)
                        merged_contacts[update_name] = merged_data
                    else:
                        # Add new contact
                        merged_contacts[update_name] = update_data
            else:
                # No update file - just process source contacts
                self.logger.info("No update file specified - processing source contacts only")
                merged_contacts = source_contacts.copy()
                self.logger.info(f"Processing {len(merged_contacts)} source contacts")
            
            # Remove duplicates if requested
            if remove_duplicates_flag:
                merged_contacts = self.processor.remove_duplicates(merged_contacts)
                # Update output filename to indicate duplicates were removed
                base_name = os.path.splitext(output_file)[0]
                output_file = f"{base_name}_no_duplicates.vcf"
            
            # Write merged contacts
            self.processor.write_vcf(merged_contacts, output_file)
            
            self.logger.info("Merging process completed successfully")
            return output_file
            
        except Exception as e:
            self.logger.error(f"Error during VCF update: {e}")
            raise

    def validate_configuration(self) -> bool:
        """
        Validates the configuration file.
        Returns True if valid, False otherwise.
        """
        config = self.config.config
        if not config:
            self.logger.error("Configuration not loaded. Cannot validate.")
            return False

        # Check input files
        source_file = config.get('input_files', {}).get('source')
        update_file = config.get('input_files', {}).get('update')
        output_file = config.get('output_file')

        if not source_file:
            self.logger.error("Source file path not specified in config.")
            return False
        if not os.path.exists(source_file):
            self.logger.error(f"Source file not found at: {source_file}")
            return False

        if update_file and not os.path.exists(update_file):
            self.logger.warning(f"Update file not found at: {update_file}. Will process only source file.")

        if not output_file:
            self.logger.error("Output file path not specified in config.")
            return False
        
        # No max limits enforced anymore for TEL/EMAIL; keep backward compatibility if present.
        
        # Check log level
        log_level = config.get('log_level', 'INFO')
        if log_level.upper() not in ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']:
            self.logger.error(f"Invalid log_level '{log_level}'. Must be one of: DEBUG, INFO, WARNING, ERROR, CRITICAL.")
            return False
        
        self.logger.info("Configuration validated successfully.")
        return True

    def get_processing_stats(self) -> Dict[str, Any]:
        """
        Returns statistics about the last processing run.
        """
        # This is a placeholder. In a real scenario, you'd store stats in a class attribute.
        # For now, we'll return dummy values.
        return {
            'total_contacts': 0,
            'contacts_processed': 0,
            'processing_time': 0.0
        }

def main() -> int:
    """Enhanced main function with better error handling and user feedback."""
    try:
        print("VCF Merger - Starting...")
        
        merger = VCFMerger()
        
        # Validate configuration before processing
        if not merger.validate_configuration():
            print("Configuration validation failed. Please check your config file.")
            return 1
        
        print("Configuration validated. Starting VCF processing...")
        output_file = merger.update_vcf_with_vcf(remove_duplicates_flag=True)
        
        # Get final statistics
        stats = merger.get_processing_stats()
        
        print("\n" + "=" * 60)
        print("PROCESSING COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        print(f"Output file: {output_file}")
        print(f"Total contacts: {stats['total_contacts']}")
        print(f"Contacts processed: {stats['contacts_processed']}")
        print(f"Processing time: {stats['processing_time']:.2f} seconds")
        print("=" * 60)
        
        return 0
        
    except KeyboardInterrupt:
        print("\nProcess interrupted by user")
        return 130
    except Exception as e:
        print(f"\nError: {e}")
        print("Check the log file 'vcf_merger.log' for detailed information")
        return 1

if __name__ == "__main__":
    exit(main())
