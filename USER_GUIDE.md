# Zone Scheduling User Guide

## Overview

Zone Scheduling is a live calendar website built on top of your Google Sheet. It helps your team:

- view school events on a calendar
- plan photographer coverage by zone
- set shared staffing limits
- add shared notes for days and zones
- search schools and review date changes
- filter the calendar by multiple scheduling attributes

The site reads its event data from the `Copy of Dates` tab in the Google Sheet and uses Apps Script to store shared planning data like limits, notes, and history.

## Main Views

### Schools

This view is for reviewing the actual school events on the calendar.

Use it to:

- see which schools are scheduled each day
- click a school to view its event details
- click a day to review all schools on that date
- review date-change history tied to that day

### Zone Staffing

This view is for staffing and capacity planning.

Use it to:

- see total photographers scheduled each day
- see photographers scheduled by zone
- compare scheduled counts against daily and zone limits
- add notes for a whole day or for a specific zone

## Data Sources

The site uses two sources of data:

### Google Sheet

The `Copy of Dates` tab provides the event data used in the calendar:

- School
- Stars
- Date
- Sent?
- Zone
- Photographers
- Type
- Confirmed

### Apps Script

Apps Script stores shared team data that should be visible across computers:

- daily limits
- zone limits
- day notes
- zone notes
- note history
- date-change history

## Calendar Colors

### Staffing warnings

- Yellow: within 2 below the limit
- Red: exactly at the limit
- Dark red: over the limit

These warnings apply to:

- the full day in `Zone Staffing`
- each zone card inside that day

### Weekend styling

- weekends use a light purple background
- weekends with scheduled schools are highlighted red at the day level

### Zone styling

Zones use alternating grayscale accents so they are distinguishable without competing with the staffing warning colors.

## Sidebar Sections

### School Search

Search by school name to see:

- all current dates for that school
- clickable date results
- moved-date history for that school

Clicking a search result:

- switches to `Zone Staffing`
- jumps to the correct month
- scrolls to the matching day
- highlights that day

### Filter

This is the zone filter.

Use it to show:

- all zones
- or just one zone like `Z1`, `Z2`, `Z3`, `Z4`, or `Z5`

### Filter By Attribute

This lets you filter the calendar by:

- Star Ranking
- Event Type
- Sent
- Confirmed

These filters can be combined with the zone filter.

Examples:

- `Z3` + `Seniors`
- `Sent` + `Confirmed`
- `Not confirmed` + `Underclass`

When filters are active:

- a `Filtered View` banner appears above the calendar
- the calendar background changes to light red

### Details

The `Details` panel changes based on what you click.

In `Schools`:

- clicking a school shows event details
- clicking a day shows all schools on that day plus date-change history

In `Zone Staffing`:

- clicking a day shows which schools make up that day’s staffing totals

### Day Notes

The notes panel supports:

- day notes
- zone notes
- note history
- holiday display
- current note save timestamp

These notes are shared across computers.

## Notes

### Day notes

Click the day number on the calendar to open a note for that date.

### Zone notes

In `Zone Staffing`, each zone has a `Note` button.

Clicking that opens a note tied to that specific:

- day
- zone

### Note history

If a note is replaced or cleared, the previous note is kept in history and shown in the notes panel.

## Holidays

The site includes built-in holidays and shows them:

- as a chip in the calendar day
- in the notes panel as read-only holiday context

Included holidays currently cover:

- US federal holidays
- Rosh Hashanah
- Yom Kippur
- Lunar New Year
- Eid al-Fitr
- Diwali

## Staffing Limits

In `Zone Staffing`, you can manually set:

- a `Daily limit`
- a `Limit` for each zone

These limits:

- are shared across computers
- are saved through Apps Script
- control the day and zone warning colors

## Confirmed Tracking

The `Confirmed` attribute filter uses the `Confirmed` field from `Copy of Dates`, which is pulled from the checkbox in `Table Main Tab` column `W`.

This gives you a lightweight way to track which entries have already been added to Roosted without connecting the Roosted API yet.

## Date History

The site also supports date-change history.

When dates are changed in `Table Main Tab`, Apps Script tracks them and the website can show:

- moved-date history in school search
- day-level change history in the `Details` panel

This is especially useful when multiple people are editing the source sheet.

## Typical Workflow

One common way to use the site:

1. Open `Zone Staffing`
2. Jump to the current month
3. Review daily and zone totals
4. Set or adjust daily and zone limits
5. Check for yellow, red, and dark red warning days
6. Click a day to see which schools are driving the counts
7. Add a day note or zone note if the team needs context
8. Use `Confirmed` to filter what has already been added to Roosted
9. Use `School Search` to check a school’s dates and change history

## Important Notes

- The Google Sheet remains the source of truth for event data
- Limits and notes are shared through Apps Script
- The site does not currently write event changes back to the sheet
- The site does not currently connect directly to RoostedHR
- The `Confirmed` filter is a manual bridge for tracking what has been entered into Roosted

## Project Files

- `index.html`: page structure and UI sections
- `styles.css`: colors, layout, and visual styling
- `app.js`: data loading, filtering, calendar behavior, notes, and shared logic
