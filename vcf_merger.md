# VCF Merger Tool

## 📌 Zweck
Das **VCF Merger Tool** dient zur Verarbeitung, Zusammenführung und Bereinigung von vCard-Dateien (`.vcf`).  
Es unterstützt insbesondere iCloud-Exporte, bei denen E-Mail-Adressen teilweise fälschlich in `NOTE`-Feldern gespeichert werden.

### Features
- Einlesen einer **Quell-VCF** und optional einer **Update-VCF**
- **Zusammenführung** mit automatischer Konfliktauflösung
- **Extraktion von E-Mails** aus `NOTE`-Feldern (kritisch für iCloud)
- **Dublettenerkennung** mit Auswahl des vollständigsten Kontakts
- **Backup** von bestehenden Dateien
- **Logging** in Datei + Konsole

---

## ⚙️ Eingaben & Ausgaben

### Eingaben
- **Konfigurationsdatei**: `vcf_config.json`
- **VCF-Dateien**:
  - Quell-VCF (`source`)
  - Update-VCF (`update`, optional)

### Ausgaben
- **Finale VCF-Datei**: `contacts_final.vcf`  
- Bei Dublettenbereinigung: `contacts_final_no_duplicates.vcf`
- **Backup-Dateien**: `contacts_final.vcf_backup_YYYYMMDD_HHMMSS`
- **Logdatei**: `vcf_merger.log`

---

## 🏗 Hauptkomponenten

### 1. Konfigurationsmanagement (`VCFConfig`)
- Laden & Speichern der Konfigurationsdatei
- Standardwerte für Pfade, Limits, Logging

### 2. Logging (`setup_logging`)
- Logausgabe in Datei + Konsole
- Level: DEBUG, INFO, WARNING, ERROR, CRITICAL

### 3. Parsing (`VCFParser`)
- **Namefelder (FN, N)**
- **Geburtsdaten (BDAY)** → Normalisierung auf `YYYY-MM-DD`
- **Telefonnummern (TEL)** → verschiedene Formate, Validierung & Deduplizierung
- **E-Mails (EMAIL)** → Standard + Extraktion aus `NOTE`-Feldern
- **Adressen (ADR)** → No
