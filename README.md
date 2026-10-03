# Lucky Cement Table Tennis Tournament

A single-page tournament dashboard (`tournament-dashboard.html`) published as a claude.ai Artifact
with a shared live database (`db` + `user` capabilities).

Live page: https://claude.ai/artifact/5MTdXjNtHkRhYL9SSVwMuL

## What players can do
- **Register** for Singles, Doubles or both (name + optional department).
- **Pick a doubles partner** from players who are still free. The choice is final: the pair is locked and
  only the organizer can remove or change it.
- **Book a match**: choose category, opponent, day and an open slot between 4:30 and 7:30 pm.
  Slot/table and player double-booking are both prevented.
- **Report scores** game by game once the match has started (games to 11, win by 2). The opponent
  confirms or disputes the result. Standings update automatically.

## What the organizer (artifact owner) can do — Admin tab
- Set match days, number of tables, best of 3/5, and 20- or 30-minute slots.
- Add players who won't sign in, toggle categories, remove players.
- Create and remove doubles pairs.
- Delete any match, enter or correct any result (organizer entries are final), clear results.

## Data model
| Path | Contents | Who writes |
| --- | --- | --- |
| `config/settings` | days, tables, bestOf, slotMinutes | owner only |
| `players/{id}` | name, dept, singles, doubles, uid | the player (own doc) or owner |
| `pairs/{id}` | a, b (player ids) | the player who chose (own doc) or owner |
| `matches/{day_hhmm_tN}` | category, a, b, day, slot, table, status, games, winner | participants |
| `locks/pairing` | lease used to stop two people grabbing the same partner | anyone |

Match status flow: `scheduled` → `reported` → `confirmed` (or `disputed`).

## Sharing
Players need write access to register: on a Team/Enterprise plan set general access to
**Contributor**; on a personal plan invite each player by email as an **Editor**.
