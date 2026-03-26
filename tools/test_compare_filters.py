"""Tests for compare dialog filter and search functionality.

Tests the filtering logic used in _toner_open_compare_dialog() — dropdown
filters (type, make, model, player, mouthpiece) and text search across
all profile fields.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))


# ---------------------------------------------------------------------------
# Filter logic extracted from _toner_open_compare_dialog.refresh_list()
# ---------------------------------------------------------------------------

def filter_profiles(all_profiles, filter_type="All", filter_make="All",
                    filter_model="All", filter_player="All",
                    filter_mpc="All", search=""):
    """Apply dropdown filters and text search to a list of profiles.

    Args:
        all_profiles: list of (lib_name, prof_name, prof_data) tuples
        filter_type/make/model/player/mpc: dropdown values ("All" = no filter)
        search: text search string (case-insensitive substring)

    Returns:
        list of (lib_name, prof_name, prof_data) tuples that pass all filters
    """
    results = []
    search = search.strip().lower()
    for lib_name, prof_name, prof in all_profiles:
        if filter_type != "All" and prof.get('horn_type', '') != filter_type:
            continue
        if filter_make != "All" and prof.get('horn_make', '') != filter_make:
            continue
        if filter_model != "All" and prof.get('horn_model', '') != filter_model:
            continue
        if filter_player != "All" and prof.get('player', '') != filter_player:
            continue
        if filter_mpc != "All" and prof.get('mouthpiece', '') != filter_mpc:
            continue
        if search:
            haystack = " ".join([
                prof_name,
                prof.get('horn_make', ''),
                prof.get('horn_model', ''),
                prof.get('serial', ''),
                prof.get('player', ''),
                prof.get('mouthpiece', ''),
                prof.get('reed', ''),
                prof.get('notes', ''),
            ]).lower()
            if search not in haystack:
                continue
        # Must have captures (mirrors refresh_list behavior)
        sessions = [s for s in prof.get('sessions', []) if s.get('captures')]
        notes = set()
        for s in sessions:
            for c in s.get('captures', []):
                notes.add(c.get('note', ''))
        if not notes:
            continue
        results.append((lib_name, prof_name, prof))
    return results


def collect_filter_values(all_profiles):
    """Collect unique values for filter dropdowns (mirrors dialog setup)."""
    all_types = sorted(set(p.get('horn_type', '') for _, _, p in all_profiles if p.get('horn_type')))
    all_makes = sorted(set(p.get('horn_make', '') for _, _, p in all_profiles if p.get('horn_make')))
    all_models = sorted(set(p.get('horn_model', '') for _, _, p in all_profiles if p.get('horn_model')))
    all_players = sorted(set(p.get('player', '') for _, _, p in all_profiles if p.get('player')))
    all_mpcs = sorted(set(p.get('mouthpiece', '') for _, _, p in all_profiles if p.get('mouthpiece')))
    return {
        'types': all_types,
        'makes': all_makes,
        'models': all_models,
        'players': all_players,
        'mpcs': all_mpcs,
    }


# ---------------------------------------------------------------------------
# Test data — realistic profiles mirroring real app data
# ---------------------------------------------------------------------------

def _cap(note):
    """Minimal capture dict."""
    return {'note': note, 'harmonics_db': [0, -5, -10], 'fundamental_freq': 440}

PROFILES = [
    ("My Profiles", "Selmer SBA tenor 38k", {
        'horn_type': 'Tenor', 'horn_make': 'Selmer', 'horn_model': 'SBA',
        'serial': '38000', 'player': 'Matt', 'mouthpiece': 'Link STM',
        'reed': 'Vandoren 2.5', 'notes': 'piercing ring, bright',
        'sessions': [{'date': '2026-03-20', 'captures': [_cap('A4'), _cap('B4')]}],
    }),
    ("My Profiles", "Selmer BA tenor 29k", {
        'horn_type': 'Tenor', 'horn_make': 'Selmer', 'horn_model': 'BA',
        'serial': '29000', 'player': 'Matt', 'mouthpiece': 'Link STM',
        'reed': 'Vandoren 3', 'notes': 'round, full, moderate',
        'sessions': [{'date': '2026-03-20', 'captures': [_cap('A4'), _cap('C5')]}],
    }),
    ("My Profiles", "Conn NW2 Virtuoso Deluxe 205k", {
        'horn_type': 'Alto', 'horn_make': 'Conn', 'horn_model': 'Virtuoso',
        'serial': '205000', 'player': 'Matt', 'mouthpiece': 'Morgan 3C',
        'reed': 'Vandoren 2.5', 'notes': 'fat Conn tone, rich',
        'sessions': [
            {'date': '2026-03-21', 'captures': [_cap('A4'), _cap('D5')]},
            {'date': '2026-03-22', 'captures': [_cap('A4'), _cap('E5')]},
        ],
    }),
    ("My Profiles", "Couesnon Monopole II alto", {
        'horn_type': 'Alto', 'horn_make': 'Couesnon', 'horn_model': 'Monopole II',
        'serial': '12345', 'player': 'Matt', 'mouthpiece': 'Morgan 3C',
        'reed': 'Vandoren 2.5', 'notes': 'lyrical, dark, warm',
        'sessions': [{'date': '2026-03-21', 'captures': [_cap('G4')]}],
    }),
    ("My Profiles", "Selmer VI alto GP", {
        'horn_type': 'Alto', 'horn_make': 'Selmer', 'horn_model': 'Mark VI',
        'serial': '82197', 'player': 'Matt', 'mouthpiece': 'Morgan 3C',
        'reed': 'Vandoren 2.5', 'notes': 'not interesting, smooth, gold plate',
        'sessions': [{'date': '2026-03-25', 'captures': [_cap('A4'), _cap('B4')]}],
    }),
    ("My Profiles", "Keilwerth Shadow Tenor", {
        'horn_type': 'Tenor', 'horn_make': 'Keilwerth', 'horn_model': 'Shadow',
        'serial': '121000', 'player': 'Matt', 'mouthpiece': 'Link STM',
        'reed': 'Vandoren 3', 'notes': 'dark side of middle',
        'sessions': [{'date': '2026-03-20', 'captures': [_cap('A4')]}],
    }),
    ("Sample Horns", "Student YAS-62", {
        'horn_type': 'Alto', 'horn_make': 'Yamaha', 'horn_model': 'YAS-62',
        'serial': '400000', 'player': 'Student', 'mouthpiece': 'Stock 4C',
        'reed': 'Rico 2', 'notes': 'student horn, baseline',
        'sessions': [{'date': '2026-03-18', 'captures': [_cap('C5')]}],
    }),
    ("My Profiles", "Empty profile", {
        'horn_type': 'Tenor', 'horn_make': 'Selmer', 'horn_model': 'BA',
        'serial': '99999', 'player': 'Matt', 'mouthpiece': 'Link STM',
        'reed': '', 'notes': '',
        'sessions': [],  # no captures
    }),
]


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

passed = 0
failed = 0

def test(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


print("=" * 60)
print("COMPARE DIALOG FILTER TESTS")
print("=" * 60)

# --- Filter value collection ---
print("\n--- Filter value collection ---")
vals = collect_filter_values(PROFILES)
test("types includes Alto and Tenor", 'Alto' in vals['types'] and 'Tenor' in vals['types'])
test("makes includes Selmer, Conn, Couesnon, Keilwerth, Yamaha",
     all(m in vals['makes'] for m in ['Selmer', 'Conn', 'Couesnon', 'Keilwerth', 'Yamaha']))
test("models includes SBA, BA, Shadow, Virtuoso, Mark VI, Monopole II, YAS-62",
     all(m in vals['models'] for m in ['SBA', 'BA', 'Shadow', 'Virtuoso', 'Mark VI', 'Monopole II', 'YAS-62']))
test("players includes Matt and Student", 'Matt' in vals['players'] and 'Student' in vals['players'])
test("mpcs includes Link STM, Morgan 3C, Stock 4C",
     all(m in vals['mpcs'] for m in ['Link STM', 'Morgan 3C', 'Stock 4C']))
test("values are sorted", vals['makes'] == sorted(vals['makes']))

# --- No filters (all profiles with captures) ---
print("\n--- No filters applied ---")
results = filter_profiles(PROFILES)
names = [n for _, n, _ in results]
test("returns 7 profiles (excludes empty)", len(results) == 7)
test("empty profile excluded", "Empty profile" not in names)

# --- Type filter ---
print("\n--- Type filter ---")
results = filter_profiles(PROFILES, filter_type="Alto")
names = [n for _, n, _ in results]
test("type=Alto returns 4 profiles", len(results) == 4)
test("all results are altos",
     all(p.get('horn_type') == 'Alto' for _, _, p in results))
test("includes Conn, Couesnon, VI GP, YAS-62",
     all(n in names for n in ['Conn NW2 Virtuoso Deluxe 205k', 'Couesnon Monopole II alto',
                               'Selmer VI alto GP', 'Student YAS-62']))

results = filter_profiles(PROFILES, filter_type="Tenor")
test("type=Tenor returns 3 profiles", len(results) == 3)

results = filter_profiles(PROFILES, filter_type="Soprano")
test("type=Soprano returns 0", len(results) == 0)

# --- Make filter ---
print("\n--- Make filter ---")
results = filter_profiles(PROFILES, filter_make="Selmer")
names = [n for _, n, _ in results]
test("make=Selmer returns 3 profiles", len(results) == 3)
test("includes SBA, BA, VI GP",
     all(n in names for n in ['Selmer SBA tenor 38k', 'Selmer BA tenor 29k', 'Selmer VI alto GP']))

results = filter_profiles(PROFILES, filter_make="Conn")
test("make=Conn returns 1", len(results) == 1)

results = filter_profiles(PROFILES, filter_make="Buffet")
test("make=Buffet returns 0", len(results) == 0)

# --- Model filter ---
print("\n--- Model filter ---")
results = filter_profiles(PROFILES, filter_model="BA")
names = [n for _, n, _ in results]
test("model=BA returns 1 (excludes empty)", len(results) == 1)
test("result is BA tenor", names[0] == 'Selmer BA tenor 29k')

results = filter_profiles(PROFILES, filter_model="Mark VI")
test("model=Mark VI returns 1", len(results) == 1)

# --- Player filter ---
print("\n--- Player filter ---")
results = filter_profiles(PROFILES, filter_player="Matt")
test("player=Matt returns 6", len(results) == 6)

results = filter_profiles(PROFILES, filter_player="Student")
test("player=Student returns 1", len(results) == 1)

# --- Mouthpiece filter ---
print("\n--- Mouthpiece filter ---")
results = filter_profiles(PROFILES, filter_mpc="Morgan 3C")
test("mpc=Morgan 3C returns 3", len(results) == 3)

results = filter_profiles(PROFILES, filter_mpc="Link STM")
test("mpc=Link STM returns 3", len(results) == 3)

# --- Combined filters ---
print("\n--- Combined filters ---")
results = filter_profiles(PROFILES, filter_type="Alto", filter_make="Selmer")
test("type=Alto + make=Selmer returns 1 (VI GP)", len(results) == 1)
test("result is VI GP", results[0][1] == 'Selmer VI alto GP')

results = filter_profiles(PROFILES, filter_type="Tenor", filter_mpc="Link STM")
test("type=Tenor + mpc=Link returns 3", len(results) == 3)

results = filter_profiles(PROFILES, filter_type="Alto", filter_make="Selmer",
                           filter_mpc="Morgan 3C")
test("type=Alto + make=Selmer + mpc=Morgan returns 1", len(results) == 1)

results = filter_profiles(PROFILES, filter_type="Tenor", filter_make="Conn")
test("type=Tenor + make=Conn returns 0 (Conn is alto)", len(results) == 0)

# --- Text search ---
print("\n--- Text search ---")
results = filter_profiles(PROFILES, search="82197")
test("search serial '82197' returns 1 (VI GP)", len(results) == 1)
test("result is VI GP", results[0][1] == 'Selmer VI alto GP')

results = filter_profiles(PROFILES, search="205000")
test("search serial '205000' returns Conn", len(results) == 1 and results[0][1] == 'Conn NW2 Virtuoso Deluxe 205k')

results = filter_profiles(PROFILES, search="gold plate")
test("search notes 'gold plate' returns VI GP", len(results) == 1 and results[0][1] == 'Selmer VI alto GP')

results = filter_profiles(PROFILES, search="dark")
names = [n for _, n, _ in results]
test("search 'dark' returns 2 (Couesnon + Shadow)", len(results) == 2)
test("dark matches Couesnon and Shadow",
     'Couesnon Monopole II alto' in names and 'Keilwerth Shadow Tenor' in names)

results = filter_profiles(PROFILES, search="SELMER")
test("search is case-insensitive ('SELMER')", len(results) == 3)

results = filter_profiles(PROFILES, search="Morgan")
test("search mouthpiece 'Morgan' returns 3", len(results) == 3)

results = filter_profiles(PROFILES, search="Vandoren")
test("search reed 'Vandoren' returns 6", len(results) == 6)

results = filter_profiles(PROFILES, search="Rico")
test("search reed 'Rico' returns 1 (student)", len(results) == 1)

results = filter_profiles(PROFILES, search="student")
names = [n for _, n, _ in results]
test("search 'student' matches both name and notes",
     len(results) == 1 and 'Student YAS-62' in names)

results = filter_profiles(PROFILES, search="zzz_nonexistent")
test("search with no matches returns empty", len(results) == 0)

results = filter_profiles(PROFILES, search="")
test("empty search returns all (7)", len(results) == 7)

results = filter_profiles(PROFILES, search="  ")
test("whitespace-only search returns all", len(results) == 7)

# --- Combined filter + search ---
print("\n--- Combined filter + search ---")
results = filter_profiles(PROFILES, filter_type="Alto", search="Morgan")
test("type=Alto + search 'Morgan' returns 3", len(results) == 3)

results = filter_profiles(PROFILES, filter_type="Tenor", search="dark")
test("type=Tenor + search 'dark' returns 1 (Shadow)", len(results) == 1)

results = filter_profiles(PROFILES, filter_make="Selmer", search="tenor")
test("make=Selmer + search 'tenor' returns 2 (SBA, BA)", len(results) == 2)

results = filter_profiles(PROFILES, filter_type="Alto", filter_make="Conn",
                           search="fat")
test("type=Alto + make=Conn + search 'fat' returns 1", len(results) == 1)

results = filter_profiles(PROFILES, filter_type="Tenor", search="gold")
test("type=Tenor + search 'gold' returns 0 (gold plate is alto)", len(results) == 0)

# --- Edge cases ---
print("\n--- Edge cases ---")
results = filter_profiles([])
test("empty profile list returns empty", len(results) == 0)

# Profile with captures but empty fields
sparse = [("Lib", "Bare Profile", {
    'horn_type': '', 'horn_make': '', 'horn_model': '',
    'serial': '', 'player': '', 'mouthpiece': '',
    'reed': '', 'notes': '',
    'sessions': [{'date': '2026-01-01', 'captures': [_cap('A4')]}],
})]
results = filter_profiles(sparse)
test("sparse profile with captures passes no-filter", len(results) == 1)
results = filter_profiles(sparse, filter_type="Alto")
test("sparse profile excluded by type filter", len(results) == 0)
results = filter_profiles(sparse, search="bare")
test("sparse profile found by name search", len(results) == 1)

# Profile with sessions but no captures in any session
no_caps = [("Lib", "No Caps", {
    'horn_type': 'Alto', 'horn_make': 'Test', 'horn_model': 'X',
    'serial': '1', 'player': 'P', 'mouthpiece': 'M',
    'reed': '', 'notes': '',
    'sessions': [{'date': '2026-01-01', 'captures': []}],
})]
results = filter_profiles(no_caps)
test("profile with empty captures list excluded", len(results) == 0)

# Multi-session profile
multi = [("Lib", "Multi", {
    'horn_type': 'Tenor', 'horn_make': 'Selmer', 'horn_model': 'VI',
    'serial': '55555', 'player': 'Joe', 'mouthpiece': 'Meyer 5M',
    'reed': 'Legere 2.5', 'notes': 'multiple sessions test',
    'sessions': [
        {'date': '2026-01-01', 'captures': [_cap('A4')]},
        {'date': '2026-01-02', 'captures': [_cap('B4')]},
        {'date': '2026-01-03', 'captures': [_cap('C5')]},
    ],
})]
results = filter_profiles(multi, search="55555")
test("multi-session profile found by serial", len(results) == 1)
results = filter_profiles(multi, search="Legere")
test("multi-session profile found by reed", len(results) == 1)

# --- Library field in results ---
print("\n--- Library context ---")
results = filter_profiles(PROFILES, search="YAS")
test("library name preserved in results", results[0][0] == 'Sample Horns')
results = filter_profiles(PROFILES, filter_make="Selmer")
libs = set(r[0] for r in results)
test("all Selmer results from My Profiles", libs == {'My Profiles'})


# --- Summary ---
print("\n" + "=" * 60)
print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print("SOME TESTS FAILED")
print("=" * 60)

sys.exit(0 if failed == 0 else 1)
