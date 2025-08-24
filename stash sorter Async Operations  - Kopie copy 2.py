#!/usr/bin/env python3
# Path of Exile 2 Inventory Manager - Improved Version with Batch Processing & Async IO
# <3
import logging
import pyautogui
import pyperclip
import time
import threading
import keyboard
import win32gui
import os
import json
import tkinter as tk
from tkinter import ttk
import sys
import traceback
from datetime import datetime
from collections import defaultdict # Added for grouping items
import asyncio # Added for async operations

# Konfigurationsdatei
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventory_manager.log")

# Logger konfigurieren (Moved setup here for clarity)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'), # Specify encoding
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("poe2_inventory_manager") # Or your actual logger name

# Initial config placeholder (will be overwritten by load_config)
config = {}

async def copy_text_at_position(x, y):
    t_func_start = time.perf_counter()
    debug_mode = config.get("debug", {}).get("DEBUG_MODE", False)
    
    try:
        # --- Optimiert: Cache Konfigurationswerte ---
        timing_config = config.get("timing", {})
        min_duration = timing_config.get("MINIMUM_DURATION", 0.005)
        min_sleep = timing_config.get("MINIMUM_SLEEP", 0.001)
        initial_clipboard_wait = timing_config.get("CLIPBOARD_WAIT", 0.03)
        max_wait_duration = timing_config.get("CLIPBOARD_MAX_WAIT", 0.5)

        # --- Optimiert: Clipboard-Vorbereitungs-Task starten, während Maus bewegt wird ---
        # Clipboard-Task definieren
        async def prepare_clipboard():
            try:
                initial_content = await asyncio.to_thread(pyperclip.paste)
                await asyncio.to_thread(pyperclip.copy, "")
                # Schnellerer Check ohne while-Schleife
                await asyncio.sleep(0.03)  # Erhöht von 0.02s auf 0.03s
                check = await asyncio.to_thread(pyperclip.paste)
                if check != "":
                    await asyncio.to_thread(pyperclip.copy, "")  # Ein zweiter Versuch
                    await asyncio.sleep(0.02)  # Erhöht von 0.01s auf 0.02s
                return True, initial_content
            except Exception as e:
                logger.warning(f"Slot ({x},{y}): Clipboard prepare error: {e}")
                return False, "<CLIPBOARD_ERROR>"
                
        # Beide Tasks parallel starten
        clipboard_task = asyncio.create_task(prepare_clipboard())
        
        # Maus bewegen (während Clipboard vorbereitet wird)
        await asyncio.to_thread(pyautogui.moveTo, x, y, duration=min_duration)
        
        # ANPASSUNG: Nochmals erhöhte Verzögerung nach Mausbewegung
        await asyncio.sleep(0.09)  # Von 0.07s auf 0.09s erhöht für bessere Verlässlichkeit
        
        # Warten auf Clipboard-Vorbereitung
        clear_success, initial_clipboard_content = await clipboard_task
        
        if not clear_success:
            logger.error(f"Slot ({x},{y}): Clipboard clearing failed.")
            return ""

        # --- Optimiert: Copy-Befehl (nochmals erhöhte Wartezeiten) ---
        await asyncio.to_thread(lambda: (
            pyautogui.keyDown('ctrl'),
            time.sleep(0.03),  # Von 0.025s auf 0.03s erhöht
            pyautogui.press('c'),
            time.sleep(0.03),  # Von 0.025s auf 0.03s erhöht
            pyautogui.keyUp('ctrl')
        ))

        # --- ANPASSUNG: Längere Wartezeit für die Spiel-Engine ---
        await asyncio.sleep(0.05)  # Von 0.04s auf 0.05s erhöht
        
        # --- Optimiert: Clipboard-Abruf mit Early Return und schnellerem Timeout für leere Slots ---
        start_wait = time.perf_counter()
        paste_attempts = 0
        current_wait_interval = initial_clipboard_wait
        
        # Sofortiger erster Check
        current_clipboard = await asyncio.to_thread(pyperclip.paste)
        
        # Wenn leer, einen zweiten Check sofort versuchen, anstatt lange zu warten
        if not current_clipboard or current_clipboard == initial_clipboard_content:
            await asyncio.sleep(0.06)  # Eine kurze Pause für einen zweiten Versuch, von 0.05s auf 0.06s erhöht
            current_clipboard = await asyncio.to_thread(pyperclip.paste)
            
            # Wenn nach zweitem Check immer noch leer, ist der Slot wahrscheinlich leer
            # Früher zurückkehren mit leerem String
            if not current_clipboard or current_clipboard == initial_clipboard_content:
                if debug_mode:
                    logger.debug(f"Slot ({x},{y}): Wahrscheinlich leer (schnelles Return nach {time.perf_counter() - t_func_start:.4f}s)")
                return ""
        
        # Wenn der Slot nicht leer ist (erfolgreicher früher Check)
        if current_clipboard and current_clipboard != initial_clipboard_content:
            if debug_mode:
                logger.debug(f"Slot ({x},{y}): Quick success after {time.perf_counter() - start_wait:.4f}s")
            
            # ANPASSUNG: Mehr Zeit für item-Text verarbeitung
            await asyncio.sleep(0.03)  # Von 0.02s auf 0.03s erhöht
            return current_clipboard
            
        # Weiteres Warten falls nötig (seltenere Fälle)
        while time.perf_counter() - start_wait < max_wait_duration:
            paste_attempts += 1
            await asyncio.sleep(current_wait_interval)
            
            current_clipboard = await asyncio.to_thread(pyperclip.paste)
            if current_clipboard and current_clipboard != initial_clipboard_content:
                break
                
            # Optimiert: Langsamere Steigerung der Wartezeit
            current_wait_interval = min(current_wait_interval * 1.2, 0.1)
        
        # Debug-Informationen nur wenn nötig
        if debug_mode and current_clipboard:
            first_line = current_clipboard.splitlines()[0] if '\n' in current_clipboard else current_clipboard
            log_text = (first_line[:50] + '...') if len(first_line) > 50 else first_line
            logger.debug(f"Slot ({x},{y}): '{log_text}', {paste_attempts+1} attempts, {time.perf_counter() - t_func_start:.4f}s")
            
        return current_clipboard or ""

    except Exception as e:
        logger.error(f"Slot ({x},{y}): Error in copy_text: {e}")
        try:
            await asyncio.to_thread(pyautogui.keyUp, 'ctrl')
        except Exception: pass
        return ""

# --- Item Name Listen ---
CHANCE_BASE_TYPES = [
    "stellar amulet", "sapphire ring", "emerald ring",
    "ornate belt", "gold ring", "gold amulet", "heavy belt",
    "solar amulet"
]
ULTIMATUM_DJINN_NAMES = ["inscribed ultimatum", "djinn barya"]
# ------------------------

# Default-Konfiguration
DEFAULT_CONFIG = {
    "timing": {
        "MINIMUM_DURATION": 0.005, "MINIMUM_SLEEP": 0.001, "PAUSE": 0.005,
        "DARWIN_CATCH_UP_TIME": 0, "WINDOW_CHECK_INTERVAL": 0.5,
        "CLIPBOARD_WAIT": 0.1, "TAB_SWITCH_WAIT": 0.3, "POST_CLICK_WAIT": 0.1
    },
    "inventory": {
        "ROWS": 5, "COLUMNS": 12, "FIRST_SLOT_TOP_LEFT_X": 1600,
        "FIRST_SLOT_TOP_LEFT_Y": 876, "SLOT_WIDTH": 75, "SLOT_HEIGHT": 75
    },
    "stash_tabs": {
        "RARE": {"X": 1169, "Y": 153},
        "RUNE": {"X": 1125, "Y": 1047},
        "JEWEL": {"X": 1123, "Y": 714},
        "QUALITY_SOCKET": {"X": 1146, "Y": 745},
        "PRECURSOR_TABLET": {"X": 1193, "Y": 223}, # Also used for Omens now
        "CHANCE_ITEMS": {"X": 0, "Y": 0},         # Must be calibrated
        "ULTIMATUM_DJINN": {"X": 0, "Y": 0},      # Must be calibrated
        "CURRENCY_CATALYST": {"X": 0, "Y": 0}     # Must be calibrated (for stackable currency like catalysts)
        # Add "VENDOR_CHAOS": {"X": 0, "Y": 0} if using chaos recipe feature later
    },
    "game": { "WINDOW_TITLE": "Path of Exile" }, # Adjust title if needed
    "debug": { "DEBUG_MODE": True, "PROGRESSIVE_SCAN": True },
    "active_profile": "default",
    "profiles": {} # Profiles stored here
}

# --- Global Variables ---
running = False
status_window = None
status_label = None
profile_var = None
# item_texts = [] # Seems unused, commented out
ALL_COORDINATES = []
last_window_check_time = 0
last_window_check_result = False
slots_found_empty_or_ignored = set() # Correct global variable for progressive scan
overlay_window = None                # NEW: For grid overlay
overlay_canvas = None                # NEW: For grid overlay
overlay_visible = False              # NEW: For grid overlay
last_mouse_pos = None                # NEW: Cache for last mouse position
_item_pattern_cache = {}             # NEW: Cache for item pattern matching
_item_decision_cache = {}            # NEW: Cache for item decisions

# --- Funktionen load_config bis calibrate_stash_tab ---
def load_config():
    """Loads configuration from JSON file, merging with defaults."""
    global config # Allow modification of the global config dict
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
            config = DEFAULT_CONFIG.copy() # Start fresh with defaults

            def update_dict_recursively(d, u): # Recursive update function
                for k, v in u.items():
                    if isinstance(v, dict) and k in d and isinstance(d.get(k), dict):
                         d[k] = update_dict_recursively(d.get(k, {}).copy(), v)
                    else:
                        d[k] = v
                return d

            config = update_dict_recursively(config, loaded_config) # Merge loaded into defaults
            logger.info(f"Konfiguration aus {CONFIG_FILE} geladen und mit Defaults gemischt.")
        else:
            config = DEFAULT_CONFIG.copy()
            save_config() # Create file with defaults if it doesn't exist
            logger.info(f"Default-Konfiguration erstellt und in {CONFIG_FILE} gespeichert.")

    except json.JSONDecodeError as e:
        logger.error(f"Fehler beim Parsen der Konfigurationsdatei {CONFIG_FILE}: {e}")
        logger.info("Verwende Default-Konfiguration.")
        config = DEFAULT_CONFIG.copy()
    except Exception as e:
        logger.error(f"Unerwarteter Fehler beim Laden der Konfiguration: {e}", exc_info=True)
        logger.info("Verwende Default-Konfiguration.")
        config = DEFAULT_CONFIG.copy()

    # Apply timing settings to pyautogui AFTER config is loaded/defaulted
    pyautogui.MINIMUM_DURATION = config.get("timing", {}).get("MINIMUM_DURATION", 0.005)
    pyautogui.MINIMUM_SLEEP = config.get("timing", {}).get("MINIMUM_SLEEP", 0.001)
    pyautogui.PAUSE = config.get("timing", {}).get("PAUSE", 0.005)
    pyautogui.DARWIN_CATCH_UP_TIME = config.get("timing", {}).get("DARWIN_CATCH_UP_TIME", 0)

    # Ensure essential keys exist after loading potentially incomplete config
    config.setdefault("inventory", DEFAULT_CONFIG["inventory"])
    config.setdefault("stash_tabs", DEFAULT_CONFIG["stash_tabs"])
    config.setdefault("game", DEFAULT_CONFIG["game"])
    config.setdefault("debug", DEFAULT_CONFIG["debug"])
    config.setdefault("timing", DEFAULT_CONFIG["timing"])
    config.setdefault("profiles", {})
    config.setdefault("active_profile", "default")

    # Load active profile if specified and exists
    active_profile_name = config.get("active_profile", "default")
    if active_profile_name != "default":
        if not load_profile(active_profile_name): # load_profile handles logging
             config["active_profile"] = "default" # Fallback to default if load fails
             logger.warning(f"Konnte aktives Profil '{active_profile_name}' nicht laden, verwende 'default'.")
        else:
             logger.info(f"Aktives Profil '{active_profile_name}' beim Start geladen.")
    else:
        # Ensure defaults are loaded if active profile is "default"
        load_profile("default")


    precalculate_coordinates() # Recalculate coords after loading config/profile


def save_config():
    """Saves the current configuration state to the JSON file."""
    global config
    try:
        # Create a copy to save
        config_to_save = config.copy()
        # Ensure profiles key exists
        config_to_save.setdefault("profiles", {})

        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_to_save, f, indent=4, sort_keys=True)
        logger.info(f"Konfiguration erfolgreich in {CONFIG_FILE} gespeichert.")
    except Exception as e:
        logger.error(f"Fehler beim Speichern der Konfiguration: {e}", exc_info=True)


def save_profile(profile_name):
    """Saves the current inventory and stash tab settings as a named profile."""
    global config
    if not profile_name or profile_name.strip() == "" or profile_name == "default":
        logger.warning("Speichern fehlgeschlagen: Profilname darf nicht leer oder 'default' sein.")
        return False
    profile_name = profile_name.strip()

    # Ensure profiles dictionary exists
    config.setdefault("profiles", {})

    # Get current settings to save into the profile
    current_inventory = config.get("inventory", DEFAULT_CONFIG["inventory"]).copy()
    current_stash_tabs = config.get("stash_tabs", DEFAULT_CONFIG["stash_tabs"]).copy()

    config["profiles"][profile_name] = {
        "inventory": current_inventory,
        "stash_tabs": current_stash_tabs
    }
    config["active_profile"] = profile_name # Set the newly saved profile as active
    save_config() # Save the entire config file with the new profile
    logger.info(f"Profil '{profile_name}' gespeichert und als aktiv gesetzt.")
    return True


def load_profile(profile_name):
    """Loads inventory and stash tab settings from a named profile."""
    global config
    if not profile_name or profile_name == "default":
        # Load defaults directly into current config
        config["inventory"] = DEFAULT_CONFIG["inventory"].copy()
        config["stash_tabs"] = DEFAULT_CONFIG["stash_tabs"].copy()
        config["active_profile"] = "default"
        precalculate_coordinates()
        update_overlay_grid() # !!! Update overlay after loading default profile !!!
        logger.info("Default-Profil ('default') geladen.")
        return True

    if "profiles" in config and profile_name in config["profiles"]:
        profile = config["profiles"][profile_name]
        # Load profile settings, falling back to defaults if keys are missing in the profile
        config["inventory"] = profile.get("inventory", DEFAULT_CONFIG["inventory"]).copy()
        # Merge profile stash tabs with defaults (profile overrides defaults)
        merged_tabs = DEFAULT_CONFIG["stash_tabs"].copy()
        merged_tabs.update(profile.get("stash_tabs", {}))
        config["stash_tabs"] = merged_tabs

        config["active_profile"] = profile_name
        precalculate_coordinates()
        update_overlay_grid() # !!! Update overlay after loading named profile !!!
        logger.info(f"Profil '{profile_name}' erfolgreich geladen.")
        return True
    else:
        logger.warning(f"Profil '{profile_name}' nicht in der Konfiguration gefunden.")
        return False


def precalculate_coordinates():
    """Calculates and caches the center coordinates of all inventory slots."""
    global ALL_COORDINATES, config # Need config
    coordinates = []
    try:
        inv_config = config.get("inventory", {})
        rows = inv_config.get("ROWS")
        cols = inv_config.get("COLUMNS")
        start_x = inv_config.get("FIRST_SLOT_TOP_LEFT_X")
        start_y = inv_config.get("FIRST_SLOT_TOP_LEFT_Y")
        slot_w = inv_config.get("SLOT_WIDTH")
        slot_h = inv_config.get("SLOT_HEIGHT")

        if not all([isinstance(rows, int), isinstance(cols, int),
                    isinstance(start_x, int), isinstance(start_y, int),
                    isinstance(slot_w, int), isinstance(slot_h, int),
                    rows > 0, cols > 0, slot_w > 0, slot_h > 0]):
            raise ValueError("Ungültige oder fehlende Inventar-Konfigurationswerte.")

        half_w = slot_w // 2
        half_h = slot_h // 2

        for r in range(rows):
            for c in range(cols):
                center_x = start_x + (c * slot_w) + half_w
                center_y = start_y + (r * slot_h) + half_h
                coordinates.append((center_x, center_y))

        ALL_COORDINATES = coordinates
        logger.debug(f"{len(coordinates)} Slot-Koordinaten berechnet ({rows}x{cols}). Start: ({start_x},{start_y}), Size: ({slot_w}x{slot_h})")
        if not coordinates:
             logger.warning("Koordinatenberechnung ergab eine leere Liste. Bitte Inventar kalibrieren.")
    except (KeyError, TypeError, ValueError) as e:
        logger.error(f"Fehler bei der Koordinatenberechnung: {e}. Bitte Inventar-Einstellungen prüfen/kalibrieren.")
        ALL_COORDINATES = []
    except Exception as e:
        logger.error(f"Unerwarteter Fehler bei Koordinatenberechnung: {e}", exc_info=True)
        ALL_COORDINATES = []


def is_game_window_active_sync():
    """Checks if the Path of Exile window is currently the foreground window. (SYNCHRONOUS)"""
    global last_window_check_time, last_window_check_result, config # Need config
    current_time = time.time()
    check_interval = config.get("timing", {}).get("WINDOW_CHECK_INTERVAL", 0.5)

    if current_time - last_window_check_time < check_interval:
        return last_window_check_result

    try:
        active_window_handle = win32gui.GetForegroundWindow()
        if active_window_handle:
            active_window_title = win32gui.GetWindowText(active_window_handle)
            target_title = config.get("game", {}).get("WINDOW_TITLE", "Path of Exile")
            result = target_title.lower() in active_window_title.lower()
        else:
            result = False
    except Exception as e:
        if current_time - last_window_check_time > 5:
             logger.warning(f"Fehler beim Prüfen des aktiven Fensters: {e}", exc_info=False)
        result = False

    last_window_check_time = current_time
    last_window_check_result = result
    return result

async def is_game_window_active_async():
    """Async wrapper for checking the game window status."""
    return await asyncio.to_thread(is_game_window_active_sync)


def update_status(message, color="black"):
    """Updates the status label in the GUI thread-safely."""
    if status_label and status_window:
        try:
            if status_window.winfo_exists():
                status_window.after(0, lambda: status_label.config(text=message, foreground=color))
        except tk.TclError as e:
             if "invalid command name" not in str(e).lower():
                  logger.warning(f"GUI Status Update Fehler: {e}")
        except Exception as e:
            logger.error(f"Unerwarteter GUI Status Update Fehler: {e}", exc_info=True)


async def move_mouse_and_click(x, y, ctrl_click=False):
    """Optimierte Version mit adaptiven Wartezeiten und Positionscache"""
    global config, last_mouse_pos
    
    try:
        # Timing aus Config abrufen (nur einmal)
        timing_config = config.get("timing", {})
        min_duration = timing_config.get("MINIMUM_DURATION", 0.005)
        post_click_wait = timing_config.get("POST_CLICK_WAIT", 0.1)
        
        # Optimierung: Bewegungszeit basierend auf Distanz anpassen
        move_duration = min_duration
        if last_mouse_pos:
            prev_x, prev_y = last_mouse_pos
            distance = ((x - prev_x)**2 + (y - prev_y)**2)**0.5
            # Längere Bewegungen brauchen mehr Zeit für Präzision
            if distance > 500:  # Pixeldistanz
                move_duration = max(min_duration, 0.01)  # Längere Bewegung
        
        # Maus bewegen und Position cachen
        await asyncio.to_thread(pyautogui.moveTo, x, y, duration=move_duration)
        last_mouse_pos = (x, y)
        
        # Optimierung: Ctrl-Click in einem Thread-Call
        if ctrl_click:
            # Definiere Funktion für Thread, um alle Aktionen zusammen auszuführen
            def do_ctrl_click():
                pyautogui.keyDown('ctrl')
                time.sleep(0.02)  # Reduzierte Wartezeit
                pyautogui.click()
                time.sleep(0.02)  # Reduzierte Wartezeit
                pyautogui.keyUp('ctrl')
                
            await asyncio.to_thread(do_ctrl_click)
        else:
            await asyncio.to_thread(pyautogui.click)
        
        # Optimierung: Adaptive Wartezeit nach dem Klick
        # Kürzere Wartezeit für normale Klicks, längere für Ctrl-Klicks
        wait_time = post_click_wait * (1.0 if ctrl_click else 0.7)
        await asyncio.sleep(wait_time)
        
        return True
        
    except Exception as e:
        logger.error(f"Klick-Fehler bei ({x},{y}): {e}")
        # Ctrl-Taste bei Fehler loslassen
        if ctrl_click:
            try:
                await asyncio.to_thread(pyautogui.keyUp, 'ctrl')
            except Exception:
                pass
        return False

def detect_inventory_region(): # No async needed - user interaction
    """Guides the user to calibrate inventory slot positions."""
    global config
    logger.info("Starte Inventar-Kalibrierung...")
    update_status("Kalibriere Inventar...", "blue")
    original_topmost = False
    if status_window:
        try:
             original_topmost = status_window.attributes("-topmost")
             status_window.attributes("-topmost", False)
             status_window.withdraw()
        except tk.TclError:
             original_topmost = False

    try:
        print("\n--- Inventar Kalibrierung ---")
        print("Bitte öffne dein Inventar im Spiel.")
        input("1. Bewege die Maus in die genaue MITTE des ERSTEN (obersten linken) Slots und drücke Enter...")
        x1, y1 = pyautogui.position()
        print(f"   Position 1 (Oben Links): ({x1},{y1})")

        input("2. Bewege die Maus in die genaue MITTE des ZWEITEN Slots (rechts neben dem ersten) und drücke Enter...")
        x2, y2 = pyautogui.position()
        print(f"   Position 2 (Rechts daneben): ({x2},{y2})")

        input("3. Bewege die Maus in die genaue MITTE des Slots DIREKT UNTER dem ersten Slot und drücke Enter...")
        x3, y3 = pyautogui.position()
        print(f"   Position 3 (Darunter): ({x3},{y3})")
        print("-----------------------------")

        slot_width = x2 - x1
        slot_height = y3 - y1
        first_slot_top_left_x = x1 - (slot_width // 2)
        first_slot_top_left_y = y1 - (slot_height // 2)

        if slot_width <= 0 or slot_height <= 0:
            raise ValueError(f"Ungültige Slot-Dimensionen berechnet: Breite={slot_width}, Höhe={slot_height}. Bitte erneut versuchen.")

        config.setdefault("inventory", {})
        inv_config = config["inventory"]
        inv_config["FIRST_SLOT_TOP_LEFT_X"] = first_slot_top_left_x
        inv_config["FIRST_SLOT_TOP_LEFT_Y"] = first_slot_top_left_y
        inv_config["SLOT_WIDTH"] = slot_width
        inv_config["SLOT_HEIGHT"] = slot_height
        inv_config.setdefault("ROWS", 5)
        inv_config.setdefault("COLUMNS", 12)

        save_config()
        precalculate_coordinates()
        update_overlay_grid() # !!! Update overlay after calibration !!!
        logger.info(f"Inventar erfolgreich kalibriert: Start({first_slot_top_left_x},{first_slot_top_left_y}), "
                    f"Breite={slot_width}, Höhe={slot_height}. Koordinaten neu berechnet.")
        update_status("Inventar kalibriert!", "green")
        return True

    except ValueError as e:
         logger.error(f"Inventar-Kalibrierung fehlgeschlagen: {e}")
         update_status(f"Fehler: {e}", "red")
         return False
    except Exception as e:
        logger.error(f"Unerwarteter Fehler bei Inventar-Kalibrierung: {e}", exc_info=True)
        update_status("Inventar-Kalibrierung fehlgeschlagen!", "red")
        return False
    finally:
        if status_window:
             try:
                 if not status_window.winfo_viewable():
                      status_window.deiconify()
                 status_window.attributes("-topmost", original_topmost)
             except tk.TclError:
                 pass

def calibrate_stash_tab(tab_name): # No async needed - user interaction
    """Guides the user to calibrate the position of a specific stash tab button."""
    global config
    if tab_name not in config.get("stash_tabs", {}):
        logger.error(f"Kann Tab '{tab_name}' nicht kalibrieren: Nicht in Konfiguration gefunden.")
        update_status(f"Fehler: Tab '{tab_name}' unbekannt", "red")
        return False

    logger.info(f"Starte Kalibrierung für Stash-Tab: {tab_name}...")
    update_status(f"Kalibriere '{tab_name}' Tab...", "blue")

    original_topmost = False
    if status_window:
        try:
            original_topmost = status_window.attributes("-topmost")
            status_window.attributes("-topmost", False)
            status_window.withdraw()
        except tk.TclError:
             original_topmost = False

    try:
        print(f"\n--- Stash Tab Kalibrierung: {tab_name} ---")
        print("Bitte öffne deine Stash im Spiel und stelle sicher, dass der Tab sichtbar ist.")
        input(f"1. Bewege die Maus in die genaue MITTE des '{tab_name}' Stash-Tab-Buttons und drücke Enter...")
        x, y = pyautogui.position()
        print(f"   Position für '{tab_name}': ({x},{y})")
        print("------------------------------------")

        config.setdefault("stash_tabs", {})
        config["stash_tabs"][tab_name] = {"X": x, "Y": y}
        save_config()

        logger.info(f"Stash-Tab '{tab_name}' erfolgreich kalibriert auf Position ({x}, {y}).")
        update_status(f"Tab '{tab_name}' kalibriert", "green")
        return True

    except Exception as e:
        logger.error(f"Unerwarteter Fehler bei der Kalibrierung von Tab '{tab_name}': {e}", exc_info=True)
        update_status(f"Fehler Kalibrierung '{tab_name}'!", "red")
        return False
    finally:
        if status_window:
             try:
                 if not status_window.winfo_viewable():
                      status_window.deiconify()
                 status_window.attributes("-topmost", original_topmost)
             except tk.TclError:
                 pass

# --- NEUE FUNKTIONEN für Grid Overlay --- START ---
def create_overlay_window():
    """Erstellt das transparente Overlay-Fenster für das Gitter."""
    global overlay_window, overlay_canvas, status_window

    if overlay_window and overlay_window.winfo_exists():
        try:
            overlay_window.destroy() # Zerstöre altes Fenster, falls vorhanden
        except tk.TclError: pass # Ignore error if already destroyed somehow
    overlay_window = None # Reset global state

    if not status_window or not status_window.winfo_exists():
        logger.error("Kann Overlay nicht erstellen, Hauptfenster existiert nicht.")
        return

    try:
        overlay_window = tk.Toplevel(status_window)
        overlay_window.title("Inventar Gitter Overlay")
        overlay_window.attributes("-alpha", 0.4)  # Transparenz (0.0=unsichtbar, 1.0=opak)
        overlay_window.attributes("-topmost", True) # Immer im Vordergrund
        overlay_window.overrideredirect(True) # Kein Fenstertitel/-rahmen

        # Wähle eine Farbe, die leicht transparent gemacht werden kann (oft Fuchsia)
        # Windows macht diese Farbe komplett transparent
        transparent_color = 'fuchsia' # Standard color often used for transparency
        overlay_window.attributes("-transparentcolor", transparent_color)
        overlay_window.config(bg=transparent_color) # Setze Hintergrundfarbe

        # Hole Bildschirmgröße für Vollbild-Overlay
        screen_width = overlay_window.winfo_screenwidth()
        screen_height = overlay_window.winfo_screenheight()
        overlay_window.geometry(f"{screen_width}x{screen_height}+0+0") # Vollbild

        overlay_canvas = tk.Canvas(overlay_window, bg=transparent_color, highlightthickness=0)
        overlay_canvas.pack(fill=tk.BOTH, expand=True)

        update_overlay_grid() # Zeichne das Gitter initial
        overlay_window.withdraw() # Starte versteckt
        logger.info("Overlay-Fenster erstellt.")

    except Exception as e:
        logger.error(f"Fehler beim Erstellen des Overlay-Fensters: {e}", exc_info=True)
        if overlay_window and overlay_window.winfo_exists():
             try: overlay_window.destroy() # Cleanup on error
             except tk.TclError: pass
        overlay_window = None
        overlay_canvas = None

def update_overlay_grid():
    """Zeichnet das Inventargitter auf das Overlay-Canvas basierend auf der Config."""
    global overlay_canvas, config, ALL_COORDINATES # Use globals

    if not overlay_canvas or not overlay_window or not overlay_window.winfo_exists():
        # logger.debug("Overlay Canvas nicht bereit zum Zeichnen.")
        return

    try:
        overlay_canvas.delete("grid") # Lösche alte Zeichnungen mit dem Tag "grid"

        inv_config = config.get("inventory", {})
        rows = inv_config.get("ROWS")
        cols = inv_config.get("COLUMNS")
        start_x = inv_config.get("FIRST_SLOT_TOP_LEFT_X")
        start_y = inv_config.get("FIRST_SLOT_TOP_LEFT_Y")
        slot_w = inv_config.get("SLOT_WIDTH")
        slot_h = inv_config.get("SLOT_HEIGHT")

        if not all([isinstance(rows, int), isinstance(cols, int),
                    isinstance(start_x, int), isinstance(start_y, int),
                    isinstance(slot_w, int), isinstance(slot_h, int),
                    rows > 0, cols > 0, slot_w > 0, slot_h > 0]):
            logger.warning("Ungültige Inventar-Konfig zum Zeichnen des Gitters.")
            return

        grid_color = "lime" # Leuchtende Farbe für das Gitter
        grid_width = 1       # Linienbreite

        for r in range(rows):
            for c in range(cols):
                x1 = start_x + (c * slot_w)
                y1 = start_y + (r * slot_h)
                x2 = x1 + slot_w
                y2 = y1 + slot_h
                # Zeichne Rechteck mit dem Tag "grid"
                overlay_canvas.create_rectangle(x1, y1, x2, y2, outline=grid_color, width=grid_width, tags="grid")

        # Optional: Zeichne Mittelpunkte (aus ALL_COORDINATES)
        # point_radius = 1
        # for cx, cy in ALL_COORDINATES:
        #     overlay_canvas.create_oval(cx-point_radius, cy-point_radius, cx+point_radius, cy+point_radius, fill=grid_color, outline="", tags="grid")

        # logger.debug("Overlay-Gitter neu gezeichnet.")

    except Exception as e:
        logger.error(f"Fehler beim Zeichnen des Overlay-Gitters: {e}", exc_info=True)


def toggle_overlay():
    """Schaltet die Sichtbarkeit des Overlay-Fensters um."""
    global overlay_window, overlay_visible

    # Create window if it doesn't exist or was destroyed
    if not overlay_window or not overlay_window.winfo_exists():
        create_overlay_window()
        # If creation failed, overlay_window will be None
        if not overlay_window:
             update_status("Fehler: Overlay konnte nicht erstellt werden.", "red")
             return

    try:
        if overlay_visible:
            overlay_window.withdraw() # Verstecken
            overlay_visible = False
            logger.info("Overlay ausgeblendet.")
            update_status("Overlay aus", "black")
        else:
            update_overlay_grid() # Stelle sicher, dass es aktuell ist
            overlay_window.deiconify() # Anzeigen
            overlay_window.lift() # Nach vorne bringen (just in case)
            overlay_visible = True
            logger.info("Overlay eingeblendet.")
            update_status("Overlay AN", "blue")
    except tk.TclError:
        logger.warning("Fehler beim Umschalten des Overlays (Fenster existiert nicht mehr?).")
        overlay_window = None # Reset state
        overlay_visible = False
    except Exception as e:
        logger.error(f"Fehler beim Umschalten des Overlays: {e}", exc_info=True)
# --- NEUE FUNKTIONEN für Grid Overlay --- ENDE ---


# --- GUI Erstellung (create_status_window) ---
def create_status_window():
    """Creates the Tkinter GUI window for status display and controls."""
    global status_window, status_label, profile_var, config, slots_found_empty_or_ignored # Need globals
    try:
        status_window = tk.Tk()
        status_window.title("PoE Inventarmanager (Async)")
        status_window.geometry("400x650") # Adjusted size slightly
        status_window.attributes("-topmost", True)

        style = ttk.Style()
        try:
             style.theme_use('vista') # Try a modern theme
        except tk.TclError:
             try:
                  style.theme_use('clam') # Fallback theme
             except tk.TclError:
                  logger.warning("Keine ttk-Themes gefunden, Standard-Look wird verwendet.")

        # --- Status Frame ---
        status_frame = ttk.LabelFrame(status_window, text="Status")
        status_frame.pack(padx=10, pady=(10, 5), fill=tk.X, expand=False)
        status_label = ttk.Label(status_frame, text="Initialisiere...", font=("Segoe UI", 11))
        status_label.pack(padx=10, pady=10, fill=tk.X)

        window_status_label = ttk.Label(status_frame, text="Spielfenster: Prüfe...", font=("Segoe UI", 9))
        window_status_label.pack(padx=10, pady=(0, 10), anchor=tk.W)

        # --- Profile Frame ---
        profile_frame = ttk.LabelFrame(status_window, text="Profile")
        profile_frame.pack(padx=10, pady=5, fill=tk.X, expand=False)

        profile_keys = list(config.get("profiles", {}).keys())
        profiles_list = ["default"] + sorted([p for p in profile_keys if p != "default"])
        active_profile = config.get("active_profile", "default")
        if active_profile not in profiles_list:
            active_profile = "default"
            config["active_profile"] = "default"

        profile_var = tk.StringVar(value=active_profile)

        profile_dropdown = ttk.Combobox(profile_frame, textvariable=profile_var, values=profiles_list, state="readonly")
        profile_dropdown.pack(padx=10, pady=5, fill=tk.X)

        def update_profile_dropdown():
             prof_keys = list(config.get("profiles", {}).keys())
             profs = ["default"] + sorted([p for p in prof_keys if p != "default"])
             profile_dropdown['values'] = profs
             current_active = config.get("active_profile", "default")
             if current_active not in profs:
                  profile_var.set("default")
             else:
                  profile_var.set(current_active)

        def on_profile_selected(event=None):
            selected_profile = profile_var.get()
            if selected_profile:
                success = load_profile(selected_profile)
                if success:
                    update_status(f"Profil '{selected_profile}' geladen", "green")
                else:
                    update_status(f"Fehler beim Laden von Profil '{selected_profile}'", "red")
                    profile_var.set(config.get("active_profile", "default"))
            else:
                 update_status("Kein Profil ausgewählt?", "orange")

        profile_dropdown.bind("<<ComboboxSelected>>", on_profile_selected)

        # --- New Profile Section ---
        new_profile_frame = ttk.Frame(profile_frame)
        new_profile_frame.pack(fill=tk.X, padx=10, pady=(0, 5))

        profile_entry = ttk.Entry(new_profile_frame, width=30)
        profile_entry.pack(side=tk.LEFT, pady=5, fill=tk.X, expand=True)
        profile_entry.insert(0, "Neuer Profilname")
        profile_entry.bind("<FocusIn>", lambda args: profile_entry.delete('0', 'end') if profile_entry.get() == "Neuer Profilname" else None)
        profile_entry.bind("<FocusOut>", lambda args: profile_entry.insert(0, "Neuer Profilname") if not profile_entry.get() else None)

        def save_new_profile_action():
            name = profile_entry.get().strip()
            if name and name != "Neuer Profilname" and name != "default":
                if name in config.get("profiles", {}):
                     update_status(f"Überschreibe Profil '{name}'...", "orange")
                if save_profile(name):
                     update_profile_dropdown()
                     profile_var.set(name)
                     profile_entry.delete(0, tk.END); profile_entry.insert(0, "Neuer Profilname")
                     update_status(f"Profil '{name}' gespeichert.", "green")
                else:
                     update_status(f"Fehler beim Speichern von Profil '{name}'.", "red")
            else:
                update_status("Bitte gültigen Profilnamen eingeben (nicht 'default').", "orange")

        save_profile_btn = ttk.Button(new_profile_frame, text="Speichern", command=save_new_profile_action)
        save_profile_btn.pack(side=tk.RIGHT, padx=(5, 0), pady=5)

        # --- Calibration Frame ---
        calib_frame = ttk.LabelFrame(status_window, text="Kalibrierung")
        calib_frame.pack(padx=10, pady=5, fill=tk.X, expand=False)

        calib_inventory_btn = ttk.Button(calib_frame, text="Inventar kalibrieren", command=detect_inventory_region)
        calib_inventory_btn.pack(padx=10, pady=5, fill=tk.X)

        # --- NEUER Button für Overlay ---
        overlay_toggle_btn = ttk.Button(calib_frame, text="Gitter-Overlay Umschalten", command=toggle_overlay)
        overlay_toggle_btn.pack(padx=10, pady=5, fill=tk.X)
        # --- ENDE NEUER Button ---

        # --- Stash Tab Calibration ---
        tab_calib_frame = ttk.Frame(calib_frame)
        tab_calib_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        max_cols = 2

        sorted_tabs = sorted(DEFAULT_CONFIG.get("stash_tabs", {}).keys())
        for i, tab_name in enumerate(sorted_tabs):
            btn = ttk.Button(tab_calib_frame, text=f"Tab: {tab_name}",
                             command=lambda t=tab_name: calibrate_stash_tab(t))
            row, col = divmod(i, max_cols)
            btn.grid(row=row, column=col, padx=2, pady=2, sticky="ew")
            tab_calib_frame.grid_columnconfigure(col, weight=1)

        # --- Control Frame ---
        control_frame = ttk.LabelFrame(status_window, text="Steuerung (Globale Hotkeys)")
        control_frame.pack(padx=10, pady=5, fill=tk.X, expand=False)

        start_btn = ttk.Button(control_frame, text="Start (.)", command=start_script)
        start_btn.pack(padx=10, pady=5, side=tk.LEFT, expand=True, fill=tk.X)

        stop_btn = ttk.Button(control_frame, text="Stop (Esc)", command=stop_script)
        stop_btn.pack(padx=10, pady=5, side=tk.RIGHT, expand=True, fill=tk.X)

        # --- Settings Frame ---
        settings_frame = ttk.LabelFrame(status_window, text="Einstellungen")
        settings_frame.pack(padx=10, pady=5, fill=tk.X, expand=False)

        # Debug Mode Toggle
        debug_var = tk.BooleanVar(value=config.get("debug", {}).get("DEBUG_MODE", False))
        def toggle_debug():
            is_debug = debug_var.get()
            if "debug" not in config: config["debug"] = {}
            config["debug"]["DEBUG_MODE"] = is_debug
            save_config()
            log_level = logging.DEBUG if is_debug else logging.INFO
            logger.setLevel(log_level)
            logger.info(f"Debug Modus {'aktiviert' if is_debug else 'deaktiviert'} (Level: {logging.getLevelName(log_level)}).")
            update_status(f"Debug Modus {'an' if is_debug else 'aus'}", "black")

        debug_check = ttk.Checkbutton(settings_frame, text="Debug-Modus (detaillierte Logs)", variable=debug_var, command=toggle_debug)
        debug_check.pack(padx=10, pady=5, anchor=tk.W)
        logger.setLevel(logging.DEBUG if debug_var.get() else logging.INFO)

        # Progressive Scan Toggle
        progressive_var = tk.BooleanVar(value=config.get("debug", {}).get("PROGRESSIVE_SCAN", True))
        def toggle_progressive():
            global slots_found_empty_or_ignored
            is_progressive = progressive_var.get()
            if "debug" not in config: config["debug"] = {}
            config["debug"]["PROGRESSIVE_SCAN"] = is_progressive
            save_config()
            logger.info(f"Progressives Scannen {'aktiviert' if is_progressive else 'deaktiviert'}.")
            update_status(f"ProgScan {'an' if is_progressive else 'aus'}", "black")
            slots_found_empty_or_ignored = set()
            logger.debug("Liste zu überspringender Slots (leer/ignoriert) zurückgesetzt.")

        # Corrected Checkbutton - only one needed, text updated
        progressive_check = ttk.Checkbutton(settings_frame, text="Progressives Scannen (überspringt leere/ignorierte Slots)", variable=progressive_var, command=toggle_progressive)
        progressive_check.pack(padx=10, pady=5, anchor=tk.W)


        # --- Info Label ---
        info_label = ttk.Label(status_window, text="Hotkeys: Start = Punkt (.) | Stop = Esc", font=("Segoe UI", 9))
        info_label.pack(pady=(5, 10))

        # --- Window Status Update Loop ---
        def update_window_status_display():
            if not status_window: return
            try:
                is_active = is_game_window_active_sync()
                status_text = "Spiel: Aktiv" if is_active else "Spiel: INAKTIV"
                status_color = "green" if is_active else "red"

                if status_window.winfo_exists() and window_status_label.winfo_exists():
                    window_status_label.config(text=status_text, foreground=status_color)
                if status_window.winfo_exists():
                    status_window.after(1000, update_window_status_display)
            except tk.TclError as e:
                 if "invalid command name" in str(e).lower():
                      logger.info("Statusfenster geschlossen, beende Fenster-Status-Updates.")
                 else:
                      logger.warning(f"Fehler im Fenster-Status-Update: {e}")
            except Exception as e:
                logger.error(f"Unerwarteter Fehler im Fenster-Status-Update: {e}", exc_info=True)
                if status_window and status_window.winfo_exists():
                    try:
                        status_window.after(5000, update_window_status_display)
                    except Exception: pass

        status_window.after(100, update_window_status_display)

        # --- Window Close Handler ---
        def on_close():
            logger.info("Statusfenster geschlossen (X geklickt). Skript läuft im Hintergrund weiter.")
            status_window.withdraw()

        status_window.protocol("WM_DELETE_WINDOW", on_close)
        update_status("Bereit.", "black")
        return status_window

    except Exception as e:
        logger.critical(f"Fataler Fehler beim Erstellen des Statusfensters: {e}", exc_info=True)
        return None

# ==============================================================================
# ===== ITEM IDENTIFICATION AND PROCESSING LOGIC                             =====
# ==============================================================================

def check_item_types(text):
    """Optimierte Item-Typ-Analyse mit Caching und Early Returns"""
    if not text: 
        return {'should_click': False}
        
    # Schritt 1: Hash des Textes berechnen für Cache-Lookup
    # MD5 ist schnell und für Caching ausreichend
    import hashlib
    text_hash = hashlib.md5(text.encode()).hexdigest()
    
    # Schritt 2: Cache prüfen für schnelle Antwort
    if text_hash in _item_decision_cache:
        return _item_decision_cache[text_hash]
    
    # Schritt 3: Beschleunigte Analyse
    lines = text.strip().splitlines()
    if not lines: 
        return {'should_click': False}

    # Basistyp-Dictionary erstellen
    types = {
        'precursor_tablet': False, 'jewel': False, 'rune': False, 'waystone': False,
        'tablet': False, 'flask': False, 'unique': False, 'rare': False,
        'normal': False, 'magic': False, 'quality': False, 'sockets': False,
        'currency': False, 'is_chance_base': False, 'omen': False,
        'ultimatum_djinn': False, 'stackable_currency': False,
        'should_click': False, 'first_line': lines[0].strip()
    }
    
    # Schritt 4: Early-Return-Prüfungen für schnelle Entscheidungsfindung
    # Ignorierte Basis-Typen sehr schnell prüfen (z.B. Weisheitsspruchrollen)
    first_line_lower = lines[0].lower().strip()
    IGNORE_LIST_BASES = config.get("item_definitions", {}).get("IGNORE_LIST_BASES", 
                         ["Scroll of Wisdom", "Portal Scroll"])
    
    if any(ignore_item.lower() in first_line_lower for ignore_item in IGNORE_LIST_BASES):
        _item_decision_cache[text_hash] = {'should_click': False, 'first_line': lines[0].strip()}
        return _item_decision_cache[text_hash]
    
    # Schritt 5: Textumwandlung für schnellere Verarbeitung
    text_lower = text.lower()
    
    # Schnellere Methode als vollständige Liste - schaue nur nach benötigten Zeilen
    rarity_line = None
    class_line = None
    quality_line = None
    socket_line = None
    stack_size_line = None
    
    # Ein einziger Durchlauf durch die Zeilen für alle Prüfungen
    for line in text_lower.splitlines():
        line = line.strip()
        if line.startswith("rarity:"):
            rarity_line = line
        elif line.startswith("item class:"):
            class_line = line
        elif line.startswith("quality:"):
            quality_line = line
            types['quality'] = True
        elif line.startswith("sockets:"):
            socket_line = line
            types['sockets'] = True
        elif "stack size:" in line:
            stack_size_line = line
    
    # Schritt 6: Rarität und Klasse extrahieren
    rarity = "unknown"
    if rarity_line:
        rarity = rarity_line.split(":", 1)[1].strip()
        if rarity == "unique": types['unique'] = True
        elif rarity == "rare": types['rare'] = True
        elif rarity == "magic": types['magic'] = True
        elif rarity == "normal": types['normal'] = True
        elif rarity == "currency": types['currency'] = True
    
    item_class = "unknown"
    if class_line:
        item_class = class_line.split(":", 1)[1].strip()
        if "jewel" in item_class: types['jewel'] = True
        if "flask" in item_class: types['flask'] = True
    
    # Rest der Analyselogik (ersetzt alte Methode)
    item_name_line = lines[2].lower().strip() if len(lines) > 2 else ""
    
    # Spezielle Item-Typen erkennen
    types['is_chance_base'] = any(base.lower() in text_lower for base in CHANCE_BASE_TYPES)
    types['ultimatum_djinn'] = any(name.lower() in text_lower for name in ULTIMATUM_DJINN_NAMES)
    types['precursor_tablet'] = "precursor tablet" in text_lower
    types['omen'] = "omen" in first_line_lower or "omen" in item_class
    types['waystone'] = "waystone" in first_line_lower
    types['tablet'] = "tablet" in item_class and not types['precursor_tablet']
    types['rune'] = (rarity == "currency" and "rune" in item_name_line)
    
    # Stapelbare Währungsitems erkennen
    is_stackable = (stack_size_line is not None) or \
                   ("catalyst" in first_line_lower) or \
                   ("essence" in first_line_lower) or \
                   ("oil" in first_line_lower)
    if is_stackable and rarity == "currency" and not types['rune']:
        types['stackable_currency'] = True
    
    # Allgemeine Währungsprüfung
    is_generic_currency_indicator = (rarity == "currency" or "currency" in item_class or stack_size_line is not None)
    types['currency'] = (
        is_generic_currency_indicator and
        not types['rune'] and
        not types['stackable_currency'] and
        not types['ultimatum_djinn']
    )
    
    # Schritt 7: Sollten wir klicken? Optimierte Logik
    types['should_click'] = any([
        types['precursor_tablet'], types['jewel'], types['rune'], types['waystone'],
        types['tablet'], types['flask'], types['unique'], types['rare'],
        types['quality'], types['sockets'], types['currency'],
        types['is_chance_base'] and types['normal'],
        types['omen'], types['ultimatum_djinn'], types['stackable_currency']
    ])
    
    # Schritt 8: Cache-Ergebnis für zukünftige Aufrufe
    _item_decision_cache[text_hash] = types.copy()
    
    # Wenn der Cache zu groß wird, älteste Einträge entfernen
    if len(_item_decision_cache) > 1000:  # Begrenzung der Cache-Größe
        # Einfachste Methode: Cache leeren
        _item_decision_cache.clear()
    
    return types


def determine_target_destination(item_types):
    """Determines the target tab name (string) or 'AFFINITY' based on item types."""
    # This function will be replaced/modified heavily by Feature 18.2 (JSON Rules)
    # For now, it remains as it was.
    handled_by_specific_tab_or_affinity = {
        'precursor_tablet', 'omen', 'jewel', 'rune', 'ultimatum_djinn',
        'stackable_currency', 'currency'
    }
    chance_exclusions = handled_by_specific_tab_or_affinity.union({'flask', 'waystone', 'tablet'})
    rare_exclusions = handled_by_specific_tab_or_affinity
    qual_sock_exclusions = handled_by_specific_tab_or_affinity.union({'rare', 'unique'})

    if item_types.get('precursor_tablet') or item_types.get('omen'): return "PRECURSOR_TABLET"
    if item_types.get('jewel'): return "JEWEL"
    if item_types.get('rune'): return "RUNE"
    if item_types.get('ultimatum_djinn'): return "ULTIMATUM_DJINN"
    if item_types.get('stackable_currency'): return "CURRENCY_CATALYST"
    if item_types.get('currency'): return "AFFINITY"

    if item_types.get('normal') and item_types.get('is_chance_base'):
        if not any(item_types.get(k) for k in chance_exclusions):
            return "CHANCE_ITEMS"

    if item_types.get('rare') or item_types.get('unique'):
        if not any(item_types.get(k) for k in rare_exclusions):
            return "RARE"

    if item_types.get('quality') or item_types.get('sockets'):
         if not any(item_types.get(k) for k in qual_sock_exclusions):
             if not any(item_types.get(k) for k in {'flask', 'waystone', 'tablet'}):
                    return "QUALITY_SOCKET"

    if config.get("debug", {}).get("DEBUG_MODE", False):
         if item_types.get('should_click'):
             first_line = item_types.get('first_line', 'N/A')
             logger.debug(f"Item '{first_line}' ZUM KLICKEN, aber kein Ziel gefunden.")

    return None

# ==============================================================================
# ===== ASYNC PROCESSING AND CONTROL FLOW                                  =====
# ==============================================================================

async def select_stash_tab(tab_type, tab_switch_data):
    """Selects the required stash tab if not already selected, asynchronously."""
    global config
    if tab_switch_data.get("selected_tab") == tab_type:
        return True

    stash_tabs_config = config.get("stash_tabs", {})
    if tab_type not in stash_tabs_config:
        logger.error(f"Konfiguration für Stash-Tab '{tab_type}' fehlt.")
        return False

    tab_coords = stash_tabs_config[tab_type]
    try:
        tx = tab_coords.get("X")
        ty = tab_coords.get("Y")
        if not isinstance(tx, int) or not isinstance(ty, int):
             raise ValueError("Koordinaten sind keine Zahlen.")
        if tx == 0 and ty == 0:
            logger.error(f"Stash-Tab '{tab_type}' ist nicht kalibriert (Position ist 0,0).")
            return False

        if config.get("debug", {}).get("DEBUG_MODE", False):
            logger.debug(f"Wechsle zu Stash-Tab {tab_type} bei ({tx},{ty})")

        def click_tab_sync():
             min_dur = config.get("timing", {}).get("MINIMUM_DURATION", 0.005)
             pyautogui.moveTo(tx, ty, duration=min_dur)
             time.sleep(0.05)  # Increased from 0.03s to 0.05s
             pyautogui.click()

        await asyncio.to_thread(click_tab_sync)
        await asyncio.sleep(config.get("timing", {}).get("TAB_SWITCH_WAIT", 0.4))  # Increased from 0.3s to 0.4s

        tab_switch_data["selected_tab"] = tab_type
        return True

    except (KeyError, ValueError, TypeError) as e:
        logger.error(f"Fehler beim Zugriff auf Koordinaten für Tab '{tab_type}': {e}")
        return False
    except Exception as e:
        logger.error(f"Unerwarteter Fehler beim Wechsel zu Tab {tab_type}: {e}", exc_info=True)
        return False


async def process_item_queue_batched(item_queue, tab_switch_data):
    """
    Processes items in batches with parallel click operations for improved performance.
    Groups items by destination tab for efficient processing.
    """
    global running, config
    processed_slots_in_batch = set()
    debug_mode = config.get("debug", {}).get("DEBUG_MODE", False)
    
    # Helper function for active state checking..
    async def is_active():
        return running and await is_game_window_active_async()
    
    # 1. Group items by destination
    grouped_items = defaultdict(list)
    for slot_index, x, y, destination in item_queue:
        if destination:
            grouped_items[destination].append({"index": slot_index, "x": x, "y": y})
    
    # Status update with total count
    total_items = sum(len(items) for items in grouped_items.values())
    update_status(f"Verarbeite {total_items} Items...", "blue")
    
    # Batch processing function with configurable batch size
    async def process_items_in_tab(tab_name, items_list, batch_size=4):
        """Processes items in batches with configurable parallelism"""
        if not await is_active():
            return 0
            
        # Switch to tab before all operations
        if tab_name != "AFFINITY" and not await select_stash_tab(tab_name, tab_switch_data):
            logger.error(f"Tab-Wechsel zu '{tab_name}' fehlgeschlagen. Überspringe {len(items_list)} Items.")
            return 0
            
        # Add a small delay after tab switch for stability
        await asyncio.sleep(0.12)
        
        # Get the total number of slots from inventory configuration
        inv_config = config.get("inventory", {})
        num_slots = inv_config.get("ROWS", 5) * inv_config.get("COLUMNS", 12)
        
        processed = 0
        # Process items in batches
        for i in range(0, len(items_list), batch_size):
            if not await is_active():
                break
                
            batch = items_list[i:i+batch_size]
            
            # Process items in batch
            for item in batch:
                try:
                    success = await move_mouse_and_click(item["x"], item["y"], ctrl_click=True)
                    if success:
                        processed_slots_in_batch.add(item["index"])
                        processed += 1
                        
                        if debug_mode:
                            logger.debug(f"Slot {item['index']+1}: Successfully ctrl+clicked in tab {tab_name}")
                    else:
                        logger.error(f"Fehler beim Klick für Item Slot {item['index']+1}")
                    
                    # Optimized delay after each click
                    await asyncio.sleep(0.06)
                
                except Exception as e:
                    logger.error(f"Exception beim Klick für Slot {item['index']+1}: {e}")
            
            # Short pause between batches
            await asyncio.sleep(0.12)
            
        return processed
    
    # 2. Process Affinity items first
    processed_count = 0
    if "AFFINITY" in grouped_items:
        affinity_items = grouped_items.pop("AFFINITY")
        if affinity_items:
            count = len(affinity_items)
            update_status(f"Verschiebe {count} Affinity-Items...", "blue")
            processed = await process_items_in_tab("AFFINITY", affinity_items)
            processed_count += processed
            logger.info(f"{processed}/{count} Affinity-Items verarbeitet.")

    # 3. Process items for specific tabs
    # Sort tabs for consistent order
    sorted_tabs = sorted(grouped_items.keys())
    
    for target_tab in sorted_tabs:
        if not await is_active():
            break
            
        items_in_group = grouped_items[target_tab]
        if not items_in_group:
            continue
            
        count = len(items_in_group)
        update_status(f"{count} Items -> {target_tab}", "blue")
        
        # Optimized batch processing for each tab
        processed = await process_items_in_tab(target_tab, items_in_group)
        processed_count += processed
        
        logger.info(f"{processed}/{count} Items für '{target_tab}' verarbeitet.")
    
    # Final logs with total count
    total_intended = len(item_queue)
    logger.info(f"{processed_count}/{total_intended} Items erfolgreich verarbeitet.")
    
    return processed_slots_in_batch


# --- !!! THIS FUNCTION IMPLEMENTS THE IMPROVED PROGRESSIVE SCAN !!! ---
async def copy_and_process_inventory_items_async():
    """
    Scans inventory using improved progressive scan, identifies items,
    processes them, and updates the set of empty/ignored slots for the next run.
    """
    global running, slots_found_empty_or_ignored, config, ALL_COORDINATES # Need globals
    if not await is_game_window_active_async():
        logger.warning("Aktion abgebrochen: Path of Exile Fenster ist nicht aktiv.")
        update_status("Spiel nicht aktiv", "red")
        running = False
        return

    logger.info("=== Starte ASYNC Inventar Scan & Sortierung (Verbessertes ProgScan) ===")
    update_status("Scanne Inventar (Async)...", "blue")

    coords = ALL_COORDINATES
    if not coords:
        logger.error("Scan fehlgeschlagen: Inventar-Koordinaten nicht verfügbar. Bitte kalibrieren.")
        update_status("Fehler: Inventar kalibrieren!", "red")
        running = False
        return

    num_slots = len(coords)
    item_queue = []
    scan_start_time = time.time()
    debug_mode = config.get("debug", {}).get("DEBUG_MODE", False)
    progressive_scan = config.get("debug", {}).get("PROGRESSIVE_SCAN", True)

    # Set to build the list of slots to skip for the *next* run
    next_run_skips = set()

    # --- Determine Slots to Scan ---
    if progressive_scan:
        slots_to_scan_indices = [i for i in range(num_slots) if i not in slots_found_empty_or_ignored]
        if not slots_to_scan_indices:
            logger.info("Progressives Scannen: Keine neuen/änderungsbedürftigen Slots gefunden.")
            update_status("Inventar stabil oder sortiert", "green")
            running = False
            return
        logger.info(f"Progressives Scannen: Prüfe {len(slots_to_scan_indices)} von {num_slots} Slots.")
    else:
        slots_to_scan_indices = list(range(num_slots))
        logger.info(f"Scanne alle {num_slots} Slots (Progressives Scannen deaktiviert).")
        slots_found_empty_or_ignored = set()

    # --- Scan Phase ---
    items_found_for_queue = 0
    for i, slot_idx in enumerate(slots_to_scan_indices):
        if not running or not await is_game_window_active_async():
            logger.warning("Scan abgebrochen (durch Benutzer oder Fenster-Inaktivität).")
            update_status("Scan abgebrochen", "orange")
            slots_found_empty_or_ignored = next_run_skips # Update skip list before aborting
            return

        x, y = coords[slot_idx]

        if i % 5 == 0:
             progress_percent = (i + 1) / len(slots_to_scan_indices) * 100
             update_status(f"Scanne Slot {slot_idx + 1}/{num_slots} ({progress_percent:.0f}%)", "blue")

        item_text = await copy_text_at_position(x, y)

        is_empty_or_ignored = False
        if item_text:
            item_types = check_item_types(item_text) # Still using old check logic for now
            if item_types.get('should_click', False):
                target_destination = determine_target_destination(item_types) # Still using old logic
                if target_destination:
                    if debug_mode:
                         first_line = item_types.get('first_line', 'N/A')
                         log_text = (first_line[:40] + '...') if len(first_line) > 40 else first_line
                         logger.debug(f"Slot {slot_idx+1}: Item '{log_text}' -> Queue (Ziel: {target_destination})")
                    item_queue.append((slot_idx, x, y, target_destination))
                    items_found_for_queue += 1
                    is_empty_or_ignored = False # Has processable item
                else:
                    if debug_mode: logger.debug(f"Slot {slot_idx+1}: Item ZUM KLICKEN, aber KEIN ZIEL. Wird als 'ignoriert' markiert.")
                    is_empty_or_ignored = True
            else:
                if debug_mode: logger.debug(f"Slot {slot_idx+1}: Item ignoriert (should_click=False). Wird markiert.")
                is_empty_or_ignored = True
        else:
            if debug_mode: logger.debug(f"Slot {slot_idx+1}: Leer. Wird markiert.")
            is_empty_or_ignored = True

        if is_empty_or_ignored:
            next_run_skips.add(slot_idx)

    scan_duration = time.time() - scan_start_time
    logger.info(f"Async Scan Phase beendet ({scan_duration:.2f}s). {items_found_for_queue} Item(s) zur Verarbeitung vorgemerkt.")

    # --- Processing Phase ---
    if not running:
         logger.info("Verarbeitung übersprungen, da Stop-Signal während des Scans empfangen wurde.")
         update_status("Scan abgebrochen", "orange")
         slots_found_empty_or_ignored = next_run_skips
         return

    processed_slots_successfully = set()
    if item_queue:
        logger.info(f"Starte Async Verarbeitung von {len(item_queue)} Item(s)...")
        update_status(f"Verarbeite {len(item_queue)} Item(s)...", "blue")
        proc_start_time = time.time()
        tab_switch_data = {"selected_tab": None}

        processed_slots_successfully = await process_item_queue_batched(item_queue, tab_switch_data)

        proc_duration = time.time() - proc_start_time
        logger.info(f"Async Verarbeitungsphase beendet ({proc_duration:.2f}s).")
        num_processed = len(processed_slots_successfully)

        if num_processed > 0:
             update_status(f"{num_processed} Item(s) verschoben.", "green")
        elif running:
             update_status("Keine Items verschoben (Warteschlange abgearbeitet).", "orange")

    else: # Corresponds to 'if item_queue:'
        logger.info("Keine Items zum Verschieben gefunden oder vorgemerkt in dieser Runde.")
        update_status("Nichts zu verschieben", "green")

    # --- Final Step: Update Global Skip List ---
    slots_found_empty_or_ignored = next_run_skips
    if debug_mode:
        logger.debug(f"Nächster progressiver Scan wird {len(slots_found_empty_or_ignored)} Slots überspringen.")

    logger.info("=== Async Scan & Sortier Runde beendet ===")
    if running:
         update_status("Bereit für nächsten Scan.", "black")
         running = False
# --- End of copy_and_process_inventory_items_async ---


# --- Global Control Functions ---
def stop_script():
    """Signals the running process to stop."""
    global running
    if running:
        running = False
        logger.info("STOP-Signal gesendet. Aktueller Async-Vorgang wird beendet...")
        update_status("Stoppe...", "orange")
    else:
        logger.info("Stop gedrückt, aber Skript war nicht aktiv.")
        update_status("Nicht aktiv", "gray")

def start_script():
    """Starts the inventory processing in a separate thread, running the async function."""
    global running, slots_found_empty_or_ignored, config, ALL_COORDINATES
    if not running:
        # Pre-checks
        if not ALL_COORDINATES:
            logger.error("Start nicht möglich: Inventar nicht kalibriert oder Konfigurationsfehler.")
            update_status("Fehler: Inventar kalibrieren!", "red")
            return
        if not config.get("stash_tabs"):
            logger.error("Start nicht möglich: Stash-Tab Konfiguration fehlt.")
            update_status("Fehler: Stash Tabs fehlen!", "red")
            return
        if not is_game_window_active_sync():
             logger.warning("Start nicht möglich: Path of Exile Fenster nicht aktiv.")
             update_status("Spiel nicht aktiv!", "red")
             return

        # Start
        running = True
        logger.info("START-Signal empfangen (Async).")
        update_status("Starte Scan (Async)...", "blue")

        def thread_target():
            try:
                asyncio.run(copy_and_process_inventory_items_async())
            except Exception as e:
                logger.error(f"Fehler im Async Verarbeitungs-Thread: {e}", exc_info=True)
                try:
                    error_msg = str(e)[:50]
                    update_status(f"Thread Fehler: {error_msg}...", "red")
                except Exception as gui_err:
                     logger.error(f"Konnte Thread-Fehler nicht im GUI anzeigen: {gui_err}")
                global running # Need global to set running=False on error
                running = False

        thread = threading.Thread(target=thread_target, daemon=True)
        thread.start()
    else:
        logger.warning("Start gedrückt, aber ein Scan läuft bereits.")
        update_status("Läuft bereits", "orange")


# --- Exception Handling and Main Execution ---
def handle_exception(exc_type, exc_value, exc_traceback):
    """Global exception handler to log unhandled exceptions."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.critical("FATALER UNBEHANDELTER FEHLER:", exc_info=(exc_type, exc_value, exc_traceback))
    try:
         if status_window and status_label:
             error_msg = str(exc_value)
             status_msg = f"FATAL ERROR: {error_msg[:100]}" + ("..." if len(error_msg) > 100 else "")
             status_window.after(0, lambda: status_label.config(text=status_msg, foreground="red"))
    except Exception as gui_err:
        logger.error(f"Konnte fatalen Fehler nicht im GUI anzeigen: {gui_err}")


def main():
    """Main function to initialize, set up GUI and hotkeys, and run the application."""
    global config, status_window
    sys.excepthook = handle_exception

    try:
        import pyautogui, pyperclip, keyboard, win32gui, tkinter, asyncio
    except ImportError as e:
        print(f"FEHLER: Benötigte Bibliothek fehlt - {e}. "
              f"Bitte installieren (z.B. pip install pyautogui pyperclip keyboard pywin32).", file=sys.stderr)
        try:
            logging.basicConfig(level=logging.CRITICAL, filename=LOG_FILE, format='%(asctime)s - %(levelname)s - %(message)s')
            logging.critical(f"Importfehler: {e}. Installation erforderlich.")
        except Exception: pass
        sys.exit(1)

    main_start_time = time.time()
    try:
        logger.info("=============================================")
        logger.info(f"Path of Exile Inventory Manager (Async) wird gestartet...")
        logger.info(f"Datum: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Konfigurationsdatei: {CONFIG_FILE}")
        logger.info(f"Logdatei: {LOG_FILE}")
        logger.info("=============================================")

        load_config() # Load config into global 'config' variable
        if not config or not ALL_COORDINATES:
             logger.critical("Konfiguration/Koordinaten nicht geladen, Abbruch.")
             return

        status_window = create_status_window()
        if not status_window:
            logger.critical("GUI konnte nicht erstellt werden, Abbruch.")
            return

        # !!! Create overlay window AFTER main window exists !!!
        create_overlay_window()

        try:
            keyboard.add_hotkey('.', start_script, trigger_on_release=False)
            keyboard.add_hotkey('esc', stop_script, trigger_on_release=False)
            logger.info("Globale Hotkeys registriert: [.] Start Scan, [Esc] Stop Scan.")
        except Exception as e:
            logger.error(f"Fehler beim Registrieren der globalen Hotkeys: {e}. Hotkeys sind deaktiviert.", exc_info=True)
            update_status("Fehler: Hotkeys nicht aktiv!", "red")

        logger.info("Starte Tkinter Hauptschleife (Statusfenster)...")
        status_window.mainloop() # Blocks here until GUI is closed

        logger.info("Tkinter Hauptschleife beendet.")

    except KeyboardInterrupt:
        logger.info("Programm durch Benutzer (Strg+C im Terminal) beendet.")
    except Exception as e:
        logger.critical(f"Unerwarteter Fehler im Haupt-Thread: {e}", exc_info=True)
    finally:
        # --- Cleanup ---
        shutdown_start_time = time.time()
        logger.info("Beende Programm und räume auf...")

        global running
        running = False

        try:
            keyboard.unhook_all()
            logger.info("Globale Hotkeys entfernt.")
        except Exception as e:
            logger.error(f"Fehler beim Entfernen der Hotkeys: {e}", exc_info=False)

        try:
             pyautogui.keyUp('ctrl')
             pyautogui.keyUp('shift')
             pyautogui.keyUp('alt')
        except Exception: pass

        time.sleep(0.1)

        logger.info(f"Programm beendet. Laufzeit: {time.time() - main_start_time:.2f}s. Aufräumzeit: {time.time() - shutdown_start_time:.2f}s.")
        logging.shutdown()


if __name__ == "__main__":
    main()