# Workforce Planning & Rotation Automation System

> **Portfolio Disclaimer:** This project is a sanitized portfolio case study based on an Excel VBA labor planning automation I designed for operations planning. All internal system names, data sources, personnel references, and workflow specifics have been replaced with generalized equivalents. No proprietary company data, internal systems, or confidential process details are included.

---

## Project Overview

**Built with:** AI Assistant + Excel VBA  
**Purpose:** Convert a manual daily labor-planning process into a one-click macro that builds a weekly job rotation and labor board using five planning sources  
**Impact:** Reduced daily planning time by 30-60 minutes, improved assignment fairness, eliminated coverage gaps and training mismatches  
**Deployed:** April 2025 — adopted by 4 site leaders, 13+ months in active use

---

## The Problem

Operations labor planning was built manually in Excel while leaders were simultaneously managing attendance, staffing changes, safety execution, process compliance, and daily volume demands. The planner had to answer several questions at once:

"Who is available, what are they trained to do, what did they do yesterday, which work areas need coverage, how much workload needs to be handled, and how do we build a fair weekly rotation without creating repeat assignments or gaps?"

### The Manual Process (Before)

1. Pull or reference the active associate roster for the shift
2. Check time-off and scheduled absence information
3. Review trained-path eligibility to confirm which associates can perform each task
4. Check the prior-day assignment log
5. Pull work area and workload volume from a separate source
6. Manually pair each associate's second task or external task to the assigned work area
7. Avoid repeat assignments or unfair task distribution
8. Balance high-volume work areas with available trained associates
9. Create the daily labor board and repeat across the week
10. Manually mark off-days, time-off conflicts, open slots, and exception coverage
11. Update the assignment log at end of day for the next planning cycle

**Estimated time:** 30-60 minutes per planning cycle  
**Operational risk:** Planning errors, repeated assignments, unbalanced workload, missed time-off conflicts, out-of-path assignments, inconsistent rotation fairness

---

## The Solution

Used an AI assistant and Excel VBA to build a macro-driven planning tool that automates the cross-reference logic and produces a weekly job rotation and labor board. The tool:

- Excludes associates who are off, on time off, or unavailable
- Keeps associates within trained paths and tasks
- Uses prior-day assignments to reduce unnecessary repeats
- Pairs each associate's secondary task with their assigned work area
- Attaches the correct workload volume to each work area assignment
- Fills required coverage slots first, then assigns remaining roles as needed
- Flags exceptions that require leader review
- Outputs a weekly rotation board ready for shift execution

---

## How It Works

### Five Data Sources

| Source | Key Data |
|--------|----------|
| Associate Roster / Current Headcount | Who can be considered for the plan |
| Prior-Day Assignment Log / EOD Record | Reduces repeats, supports fair rotation |
| Time-Off and Absence Data | Removes unavailable associates before assignments |
| Trained-Path Eligibility Matrix | Prevents out-of-path assignment |
| Work Area / Workload Volume Data | Connects assignments to actual workload |

### Cross-Reference Engine

The main macro chains 5 dictionary lookups:

Associate ID -> Active Status -> Availability Status
             -> Training Eligibility -> Prior-Day Assignment Check
             -> Work Area Assignment -> Secondary Task Pairing
             -> Workload Volume Attached
             -> Final Assignment or Exception Flag

### Dictionary Architecture

| Dictionary | Maps |
|------------|------|
| aaToStatus | Associate to Active / Inactive / Shift Eligible |
| aaToAvailability | Associate to Available / Unavailable |
| aaToTrainedPaths | Associate to Eligible Paths and Tasks |
| aaToPriorArea | Associate to Previous Work Area |
| aaToPriorTask | Associate to Previous Secondary Task |
| areaToBagCount | Work Area to Workload Volume |
| areaToCoverageNeed | Work Area to Required Coverage |
| areaToSecondaryTask | Work Area to Secondary Task |

---

## Results and Impact

| Metric | Before | After |
|--------|--------|-------|
| Daily planning time | 30-60 min | Macro-driven, repeatable |
| Manual cross-reference errors | Multiple weekly | Eliminated |
| Coverage gaps | Frequent | Flagged before shift start |
| Training mismatches | Occasional | Prevented by eligibility check |
| Rotation fairness | Inconsistent | Systematic, prior-day aware |
| Leader onboarding time | Extended shadowing | Self-service output |

---

## AI-Assisted Development Process

Built using conversational AI with zero prior VBA experience:

1. Described the problem and 5 data sources in plain English
2. AI identified join keys and designed the data model
3. AI built the dictionary architecture and assignment logic
4. Iterated on business rules: rotation fairness, area priority, exception flagging
5. Added leader-ready output formatting and end-of-day assignment tracking

---

## Module Structure

| Module | Purpose |
|--------|---------|
| SetupLaborPlanner | Creates sheets, headers, formatting, and control buttons |
| RefreshLaborInputs | Runs all five import macros |
| BuildWeeklyRotation | Main assignment engine |
| AssignAreaCoverage | Fills required work areas with trained associates |
| PairSecondaryTask | Connects secondary tasks to work area assignments |
| ValidateRotationConflicts | Flags repeats, conflicts, and coverage gaps |
| GenerateLaborBoard | Outputs the weekly rotation board |
| ExportLeaderView | Creates clean copy for review and execution |
| FinalizeDayAssignments | Saves assignments back to assignment log for next cycle |

---

## Tools and Technologies

- AI Assistant for architecture and code generation
- Excel VBA for macro execution
- Scripting.Dictionary for in-memory hash maps
- Five internal data source integrations

---

*Built for last-mile delivery operations | Deployed April 2025 — actively in use*
