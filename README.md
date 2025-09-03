# VCF Contact Merger

Advanced VCF Contact Merger with automatic conflict resolution, duplicate removal, and comprehensive phone number parsing. Supports iCloud and Outlook formats.

## Features

- **Automatic Conflict Resolution**: Intelligently resolves conflicts between source and update data
- **Duplicate Removal**: Removes duplicate contacts, keeping the most complete entry
- **Comprehensive Phone Number Parsing**: Handles various VCF formats including iCloud and Outlook
- **Phone Number Validation**: Ensures only valid phone numbers are included
- **Email Extraction**: Extracts emails from various VCF formats
- **Address Consolidation**: Merges address information from multiple sources
- **Backup System**: Automatic backup before processing
- **Logging**: Comprehensive logging for debugging and monitoring

## Installation

1. Clone the repository:
```bash
git clone https://github.com/bgv2lr/vcf-contact-merger.git
cd vcf-contact-merger
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

Edit `vcf_config.json` to configure input/output files and advanced options:

```json
{
    "input_files": {
        "source": "source_contacts.vcf",
        "update": "update_contacts.vcf"
    },
    "output_file": "merged_contacts.vcf",
    "backup_enabled": true,
    "log_level": "INFO",
    "phone_validation": {
        "min_digits": 7,
        "check_duplicates": true,
        "allow_international": true
    },
    "conflict_resolution": {
        "auto_resolve": true,
        "prefer_update_for": ["EMAIL", "TEL", "ADR", "ORG", "NOTE"],
        "prefer_source_for": ["N", "FN", "BDAY"]
    }
}
```

## Usage

### Basic Usage
```bash
python update_private_vcf.py
```

### Advanced Usage
```python
from update_private_vcf import VCFMerger

merger = VCFMerger('config.json')
output_file = merger.update_vcf_with_vcf(remove_duplicates_flag=True)
print(f"Created: {output_file}")
```

## Conflict Resolution Rules

- **BDAY**: Prefers source unless update has default date (1900-01-01)
- **NOTE**: Prefers update (more structured information)
- **EMAIL**: Prefers update (more current)
- **TEL**: Prefers update (more current)
- **ADR**: Prefers update (more complete)
- **ORG**: Prefers update (more current)
- **N/FN**: Prefers source (better formatted)

## Phone Number Validation

- **Configurable minimum digits** (default: 7)
- **Duplicate detection** (configurable)
- **International format support**
- **iCloud and Outlook specific formats**
- **Advanced regex patterns** for complex formats
- **Comprehensive validation** with detailed logging

## File Structure

```
vcf-contact-merger/
├── update_private_vcf.py      # Main script
├── update_private_vcf_fixed.py # Fixed version
├── test_vcf_merger.py         # Unit tests
├── vcf_config.json           # Configuration
├── requirements.txt          # Dependencies
└── README.md                # This file
```

## Testing

Run the test suite:
```bash
python test_vcf_merger.py
```

## Logging

Logs are written to `vcf_merger.log` with detailed information about:
- File processing progress
- Conflict resolutions
- Phone number validations
- Duplicate removals

## Privacy

This repository contains only source code. Sensitive contact data files are excluded via `.gitignore`.

## License

Private repository - All rights reserved.

## Support

For issues and questions, please contact the repository owner.
