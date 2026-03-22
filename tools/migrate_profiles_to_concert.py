"""One-time migration: convert all tone profiles from written pitch to concert pitch."""
import sys
import os
import json

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from toner_engine import migrate_profile_to_concert

PROFILES_PATH = os.path.join(
    os.environ.get('APPDATA', ''),
    'StohrerSaxShopCompanion',
    'tone_profiles.json'
)

def main():
    with open(PROFILES_PATH, 'r') as f:
        data = json.load(f)

    migrated_count = 0
    skipped_count = 0

    for lib_name, lib_profiles in data.items():
        if not isinstance(lib_profiles, dict):
            continue
        for prof_name, prof_data in lib_profiles.items():
            if not isinstance(prof_data, dict):
                continue
            if prof_data.get('pitch_format') == 'concert':
                skipped_count += 1
                print(f"  SKIP  {lib_name} / {prof_name} (already concert)")
                continue

            horn = prof_data.get('horn_type', '?')
            n_captures = sum(
                len(s.get('captures', []))
                for s in prof_data.get('sessions', [])
            )
            new_prof = migrate_profile_to_concert(prof_data)
            data[lib_name][prof_name] = new_prof
            migrated_count += 1
            print(f"  MIGRATED  {lib_name} / {prof_name}  (horn={horn}, captures={n_captures})")

    with open(PROFILES_PATH, 'w') as f:
        json.dump(data, f, indent=2)

    print(f"\nDone: {migrated_count} migrated, {skipped_count} skipped (already concert)")

if __name__ == '__main__':
    main()
