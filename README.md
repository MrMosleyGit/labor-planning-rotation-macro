# Workforce Planning & Rotation Automation System

## Project Overview

**Built with:** Amazon Quick (AI assistant) + Excel VBA  
**Purpose:** Convert a manual daily labor-planning process into a one-click macro that builds a weekly job rotation and labor board using five planning sources  
**Impact:** Reduced daily planning time by 30-60 minutes, improved assignment fairness, eliminated coverage gaps and training mismatches  
**Deployed:** April 2025 — adopted by 4 site leaders at Amazon DLV4, 13+ months in active use

---

## The Problem

RTS labor planning was built manually in Excel while leaders were simultaneously managing attendance, staffing changes, safety execution, process compliance, and daily volume demands. The planner had to answer several questions at once:

"Who is available, what are they trained to do, what did they do yesterday, which clusters need coverage, how many bags need to be replaced, and how do we build a fair weekly rotation without creating repeat assignments or gaps?"

### The Manual Process (Before)

1. Pull or reference the active associate roster for the shift
2. Check time-off and scheduled absence information
3. Review trained-path eligibility to confirm which associates can perform each task
4. Check the prior-day assignment sheet
5. Pull cluster and bag replacement volume from a separate workload source
6. Manually pair each associate's second task or external task to the assigned cluster
7. Avoid repeat clusters, repeat external tasks, or unfair assignments
8. Balance high-bag-count clusters with available trained associates
9. Create the daily labor board and repeat across the week
10. Manually mark off-days, time-off conflicts, open slots, and exception coverage
11. Update the assignment tracker at end of day for the next cycle

**Estimated time:** 30-60 minutes per planning cycle  
**Operational risk:** Planning errors, repeated assignments, unbalanced workload, missed time-off conflicts, out-of-path assignments, inconsistent rotation fairness

---

## The Solution

Used Amazon Quick and Excel VBA to build a macro-driven planning tool that automates the cross-reference logic and produces a weekly job rotation and labor board. The tool:

- Excludes associates who are off, on time off, or unavailable
- Keeps associates within trained paths and tasks
- Uses prior-day assignments to reduce unnecessary repeats
- Pairs each associate's second task with their assigned cluster
- Attaches the correct bag replacement count to each cluster assignment
- Fills required clusters first, then assigns open slots as needed
- Flags exceptions that require leader review
- Outputs a weekly rotation board ready for shift execution

---

## How It Works

### Five Data Sources

| Source | Key Data |
|--------|----------|
| Associate Roster / Current Headcount | Who can be considered for the plan |
| Prior-Day A-Sheet / EOD Assignments | Reduces repeats, supports fair rotation |
| Time-Off Database | Removes unavailable associates before assignments |
| Trained-Path Eligibility Matrix | Prevents out-of-path assignment |
| Cluster / Bag Workload Source | Connects assignments to actual workload |

### Cross-Reference Engine

The BuildWeeklyRotation macro chains 5 dictionary lookups:

Associate ID -> Active Status -> Availability Status
             -> Training Eligibility -> Prior-Day Assignment Check
             -> Cluster Assignment -> External Task Pairing
             -> Bag Count Attached to Cluster
             -> Final Assignment or Exception Flag

### Dictionary Architecture (Hash Maps)

| Dictionary | Maps |
|------------|------|
| aaToStatus | Associate to Active/Inactive/Shift Eligible |
| aaToAvailability | Associate to Available/Unavailable |
| aaToTrainedPaths | Associate to Eligible Paths and Tasks |
| aaToPriorCluster | Associate to Previous Cluster |
| aaToPriorExternalTask | Associate to Previous Second Task |
| clusterToBagCount | Cluster to Bags to Replace |
| clusterToCoverageNeed | Cluster to Required Coverage |
| clusterToExternalTask | Cluster to Second Task |

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
4. Iterated on business rules: rotation fairness, cluster priority, exception flagging
5. Added leader-ready output formatting and end-of-day assignment tracker

---

## Module Structure

| Module | Purpose |
|--------|---------|
| SetupLaborPlanner | Creates sheets, headers, formatting, and control buttons |
| RefreshLaborInputs | Runs all five import macros |
| BuildWeeklyRotation | Main assignment engine |
| AssignClusterCoverage | Fills required clusters with trained associates |
| PairExternalTask | Connects second tasks to cluster assignments |
| ValidateRotationConflicts | Flags repeats, conflicts, and coverage gaps |
| GenerateLaborBoard | Outputs the weekly rotation board |
| ExportLeaderView | Creates clean copy for review and execution |
| FinalizeDayAssignments | Saves assignments back to A-Sheet for next cycle |

---

## Tools and Technologies

- Amazon Quick (AI assistant) for architecture and code generation
- Excel VBA for macro execution
- Scripting.Dictionary for in-memory hash maps
- Five internal data source integrations

---

*Built at Amazon Last Mile, DLV4 Delivery Station, Henderson NV | April 2025 — actively in use*
