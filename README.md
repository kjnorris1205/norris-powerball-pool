# Norris Powerball Pool

A professional-grade Microsoft Access application (.accdb) for managing Powerball lottery pool group entries, matching results, and tracking winnings. Designed for use in all 50 US states plus DC with configurable state-of-play settings.

## Requirements

- Microsoft Access 2016+ or Microsoft 365
- Single `.accdb` file (no external dependencies)
- DAO data access (Microsoft Office xx.0 Access Database Engine Object Library)

## Powerball Rules Reference

| Rule | Value |
|------|-------|
| White balls | Pick 5 from 1–69 (order does not matter) |
| Powerball | Pick 1 from 1–26 |
| Draw schedule | Monday, Wednesday, Saturday |
| Prize tiers | 9 total (0+PB through 5+PB) |

## Database Creation

1. Follow the instructions in database-setup-instructions.md

2. Follow the instructions in batch-sql-runner-instructions.md

3. Follow the instructoins in database-seeding-instructions.md