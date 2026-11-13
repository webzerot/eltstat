# Changelog

## 1.0.4
Released: 13 November 2026

- Guide: new section in Settings about program updates and the first-run setup.
- Guide: two new Troubleshooting entries ("up to date" after a push, and failed package installation).

## 1.0.3
Released: 13 November 2026

- Fixed: the first-run setup window showed "Not responding" while packages were installing. The installation now runs in the background, and the window shows a running timer.
- The installer no longer waits forever on a bad connection (network timeout and retry limits added).

## 1.0.2
Released: 13 November 2026

- First-run setup: on the very first start the program downloads requirements.txt from the GitHub repository and installs the required packages automatically.
- The setup runs once. A hidden file, .elstat_setup_done, remembers that it is finished, so later starts are instant.
- If a required package goes missing later, the setup runs again by itself.
- If GitHub has no requirements.txt or there is no internet, a built-in list of packages is used instead.
- Everything stays in the program's own folder by default, so no folders need to be changed.

## 1.0.1
- Update check test release.

## 1.0.0
- First release.
