import os
import sys
import json
import subprocess
import psutil
import re

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR_SETTINGS = get_base_path()
PROFILE_FILE = os.path.join(BASE_DIR_SETTINGS, "audio_profiles.json")
PROFILE_NAME_REGEX = re.compile(r'^[A-Za-z0-9-]+$')

class AudioProfileManager:
    def __init__(self):
        self.profiles = {}
        self.soundvolumeview_path = "SoundVolumeView.exe"
        self.config_file = PROFILE_FILE

    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            if not isinstance(config, dict) or 'profiles' not in config or not isinstance(config['profiles'], dict):
                raise ValueError("Not a valid audio_profiles.json")
            self.profiles = config.get('profiles', {})
            self.soundvolumeview_path = config.get('soundvolumeview_path', self.soundvolumeview_path)

    def apply_profile(self, profile_name):
        if profile_name not in self.profiles:
            return 0
        rules = self.profiles[profile_name]['rules']
        applied_count = 0
        for rule in rules:
            if self.execute_rule(rule):
                applied_count += 1
        return applied_count

    def execute_rule(self, rule):
        try:
            if not any(p.name().lower() == rule['app_name'].lower() for p in psutil.process_iter(['name'])):
                return False
            cmd = [
                self.soundvolumeview_path,
                '/SetAppDefault',
                rule['device_id'],
                '1',
                rule['app_name']
            ]
            subprocess.run(cmd, check=True, capture_output=True)
            return True
        except Exception:
            return False

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="SmartWindowsAudioProfiles CLI")
    parser.add_argument("profile_name", help="Profile name to activate (uses audio_profiles.json in app directory)")
    args = parser.parse_args()

    if not PROFILE_NAME_REGEX.match(args.profile_name):
        print("ERROR: Invalid profile name! Only letters, numbers, and hyphens (-) are allowed. No spaces.")
        sys.exit(1)

    app = AudioProfileManager()
    try:
        if not os.path.exists(PROFILE_FILE):
            print("ERROR: audio_profiles.json not found in the application directory.")
            sys.exit(1)
        app.load_config()
        if args.profile_name in app.profiles:
            applied = app.apply_profile(args.profile_name)
            if applied > 0:
                print(f"Profile '{args.profile_name}' activated with {applied} rule(s).")
                sys.exit(0)
            else:
                print(f"ERROR: Profile '{args.profile_name}' found but no rules could be applied. Are the target applications running? Is SoundVolumeView.exe configured correctly?")
                sys.exit(2)
        else:
            print(f"ERROR: Profile '{args.profile_name}' not found in audio_profiles.json.")
            sys.exit(1)
    except Exception as e:
        print(f"ERROR: Failed to activate profile '{args.profile_name}': {e}")
        sys.exit(2)
