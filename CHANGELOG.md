# Changelog

## 1.0.5
Released: 14 November 2026

- New: the countries that were missing from the program are now recognised: Moldavia, "ΛΟΙΠΩΝ ΕΥΡΩΠΗΣ", Monaco, Belarus, Bosnia, Uruguay, Chile, Colombia, Tunisia, Jordan, Sri Lanka, Kazakhstan, Armenia, Azerbaijan, Kyrgyzstan, Hong Kong and Afghanistan. About 90 more spellings and countries that Protel may produce later were added as well.
- New: countries can now be left out of the PDF. The country table on the report tab has a ☑ column. Click a box to leave one country out, or click the ☑ header to tick or untick all.
- Left-out countries are grey in the table. They are not in the PDF, the totals or the history, and the summary line shows how many are left out. The "PDF total differs from Excel" warning takes them into account, so it no longer appears by mistake.
- Choosing a different Excel / CSV file ticks every country again.
- Fixed: "Create PDF" error with pypdf 7. The page is now added to the new PDF first and the numbers are placed on that copy. If this fails, two other methods are tried automatically.
- The PDF is now written to a temporary file and only renamed when it is complete, so a failure never leaves a broken PDF. If every method fails, the error message names the pypdf version and the reason for each failure.
- Fixed: the program now works out where each country row sits on the form in a way that is correct on every pypdf version, so numbers cannot land on the wrong rows if a newer pypdf reports text positions differently. The output on the current templates is unchanged.
- First-run setup: pypdf is kept below version 7. A computer that already has pypdf 7 is corrected at the next start.
- First-run setup: if installing the whole list fails, each missing package is now installed one at a time, so one bad line can no longer block the rest.
- First-run setup: retries without administrator rights, installs pip if it is missing, and stops after 15 minutes instead of waiting forever.
- First-run setup: a broken or incomplete requirements.txt downloaded from GitHub is ignored and the built-in list is used instead.
- First-run setup: the details are saved in elstat_setup.log in the program folder. If the installation fails, the message shows the exact command to run by hand.

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
