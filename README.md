# POE2 Stash Sorter

Ein automatisierter Stash-Sortierer für **Path of Exile 2**, der Ihnen hilft, Ihr Inventar effizient zu organisieren und Items automatisch in die entsprechenden Stash-Tabs zu sortieren.

## 🎯 Features

- **Automatisches Item-Sortieren**: Sortiert Items automatisch basierend auf Typ und Eigenschaften
- **Multi-Profile Support**: Verschiedene Profile für unterschiedliche Auflösungen und Setups
- **Intelligente Item-Erkennung**: Erkennt Rares, Runen, Juwelen, Katalysatoren und mehr
- **Progressive Scan**: Optimiert die Scan-Geschwindigkeit durch Überspringen bereits gescannter leerer Slots
- **Async Operations**: Hochperformante asynchrone Operationen für schnellere Ausführung
- **Debug-Modus**: Ausführliche Logging-Funktionen für Troubleshooting
- **Hotkey-Support**: Einfache Steuerung über Tastenkürzel
- **Grid-Overlay**: Visueller Overlay zur Kalibrierung der Inventar-Positionen

## 🚀 Unterstützte Item-Typen

### Automatisch sortierte Items:
- **Rare Items** → Rare Tab
- **Runen** → Rune Tab  
- **Juwelen** → Jewel Tab
- **Qualitäts-Items und Socketed Items** → Quality/Socket Tab
- **Precursor Tablets & Omens** → Precursor Tab
- **Chance Base Types** → Chance Items Tab (Stellar Amulet, Sapphire Ring, etc.)
- **Ultimatum Items** → Ultimatum Tab (Inscribed Ultimatum, Djinn Barya)
- **Katalysatoren** → Currency/Catalyst Tab

### Erkannte Katalysatoren:
- Turbulent, Imbued, Intrinsic, Prismatic, Tempering, Fertile
- Accelerating, Noxious, Unstable, Abrasive, Tainted Catalyst

## 🛠️ Installation

### Voraussetzungen
- Python 3.7+
- Path of Exile 2 (Windows)
- Erforderliche Python-Bibliotheken:

```bash
pip install pyautogui pyperclip keyboard win32gui
```

### Setup
1. Repository klonen:
   ```bash
   git clone https://github.com/NotDonCitron/poe2-stash-sorter.git
   cd poe2-stash-sorter
   ```

2. Abhängigkeiten installieren:
   ```bash
   pip install -r requirements.txt
   ```

3. Konfiguration anpassen (siehe [Konfiguration](#konfiguration))

## ⚙️ Konfiguration

### Erste Einrichtung

1. **Inventar kalibrieren**: 
   - Starten Sie das Script mit `python "working mario shown.py"`
   - Verwenden Sie den Grid-Overlay (Taste `G`) zur visuellen Kalibrierung
   - Justieren Sie die Koordinaten in der Konfiguration

2. **Stash-Tabs konfigurieren**:
   - Öffnen Sie Ihre Stash in POE2
   - Notieren Sie sich die Koordinaten der Tab-Buttons
   - Aktualisieren Sie `config.json` entsprechend

### Profile

Das System unterstützt mehrere Profile für verschiedene Auflösungen:

```json
{
  "active_profile": "MEIN_PROFIL",
  "profiles": {
    "MEIN_PROFIL": {
      "inventory": {
        "FIRST_SLOT_TOP_LEFT_X": 1595,
        "FIRST_SLOT_TOP_LEFT_Y": 862,
        "SLOT_WIDTH": 79,
        "SLOT_HEIGHT": 88
      },
      "stash_tabs": {
        "RARE": {"X": 1172, "Y": 154},
        "RUNE": {"X": 1128, "Y": 1055}
      }
    }
  }
}
```

## 🎮 Verwendung

### Grundfunktionen

1. **Starten**: `python "working mario shown.py"`
2. **Hotkeys**:
   - `F8`: Start/Stop des Sortier-Prozesses
   - `F9`: Notfall-Stop (alle Operationen beenden)
   - `G`: Grid-Overlay ein-/ausschalten (zur Kalibrierung)

### Workflow

1. Öffnen Sie POE2 und gehen Sie zu Ihrem Stash
2. Öffnen Sie das Inventar
3. Starten Sie das Script
4. Drücken Sie `F8` um das automatische Sortieren zu starten
5. Das Script scannt automatisch alle Inventar-Slots und sortiert Items

## 🔧 Erweiterte Einstellungen

### Timing-Anpassungen

```json
{
  "timing": {
    "MINIMUM_DURATION": 0.003,
    "CLIPBOARD_WAIT": 0.1,
    "TAB_SWITCH_WAIT": 0.2,
    "POST_CLICK_WAIT": 0.1
  }
}
```

### Debug-Modus

Aktivieren Sie den Debug-Modus für detaillierte Logs:

```json
{
  "debug": {
    "DEBUG_MODE": true,
    "PROGRESSIVE_SCAN": true
  }
}
```

## 📝 Logs

Das Script erstellt automatisch Logs in `inventory_manager.log` mit:
- Timestamps aller Operationen
- Item-Erkennungs-Details
- Performance-Metriken
- Fehler und Warnungen

## 🔒 Sicherheit

⚠️ **Wichtige Hinweise:**
- Das Script verwendet Maus-Automatisierung und könnte als Cheat-Software erkannt werden
- Verwenden Sie es auf eigene Verantwortung
- Testen Sie immer mit weniger wertvollen Items
- Erstellen Sie Backups Ihrer wichtigen Items

## 🤝 Beitragen

Beiträge sind willkommen! Bitte:

1. Forken Sie das Repository
2. Erstellen Sie einen Feature-Branch (`git checkout -b feature/AmazingFeature`)
3. Committen Sie Ihre Änderungen (`git commit -m 'Add AmazingFeature'`)
4. Pushen Sie zum Branch (`git push origin feature/AmazingFeature`)
5. Öffnen Sie eine Pull Request

## 📄 Lizenz

Dieses Projekt steht unter der MIT-Lizenz - siehe [LICENSE](LICENSE) Datei für Details.

## ⚠️ Disclaimer

Dieses Tool ist nicht offiziell von Grinding Gear Games unterstützt. Die Nutzung erfolgt auf eigene Gefahr. Path of Exile 2 ist ein Trademark von Grinding Gear Games.

## 📞 Support

Bei Problemen oder Fragen:
- Öffnen Sie ein Issue auf GitHub
- Überprüfen Sie die Logs in `inventory_manager.log`
- Aktivieren Sie den Debug-Modus für detailliertere Informationen

---

**Happy Sorting, Exile! 🎯**