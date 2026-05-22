# JK Show and Dress Grader Agents Notes

## Project

- Repo: `git@github.com:palytinsley/jkgrader.git`
- Local path: `C:\Users\tinsl\Desktop\gas-projects\jkgrader`
- GAS Script ID: `1Wj32Xppq-IlyxQ2p--guCnTnxc3eva7ckX8DYT6NsvtVMYX4i62KWWcF`
- Spreadsheet ID: `1_xOjYXh_PQzXERVyp-k_lyge9jAxiWTPAedS6lOdl0c`
- Frontend: GitHub Pages single-file `index.html`
- Backend: `gas/Code.gs` as a Web App, execute as owner, anyone can access
- Icons: Tabler Icons via `https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css`

## Build Conventions From GLOBAL_AGENTS.md

- Use flat tab reads only. No `INDIRECT`, no dynamic range construction, no flush/sleep hacks.
- Keep one active `Code.gs`. Do not create duplicate active `doGet` or `doPost` handlers.
- All email scripts must include confirmation modal, dry run mode, and test-only send to `ktinsley@pausd.org`.
- Run `clasp status` and `clasp push` after every GAS change.
- Use descriptive git commit messages.
- Never claim deployment occurred unless the push command ran and returned success.
- Never commit `.clasprc.json`, OAuth tokens, student data, API keys, or private credentials.
- Use SSH remotes only.

## Tabs And Schemas

### `Show Night Resposnes`

Typo is real for this grader backend. Do not correct in GAS reads.

Headers:

`Timestamp | Email Address | Period | Period 5 Country Name | Period 6 Country Name | Upload 1: Your Wow factor pose | Upload 2: Front photo | Upload 3: Side photo | Upload 4: Back Photo | Program Blurb Paragraph`

### `peer-grade-responses`

`Timestamp | SubmitterName | Period | Country | Peer1Name | Peer1Percent | Peer1Justification | Peer2Name | Peer2Percent | Peer2Justification | Peer3Name | Peer3Percent | Peer3Justification | Peer4Name | Peer4Percent | Peer4Justification`

### `Roster`

`Last Name | Preferred First | Legal First | Period | Country | Email`

### `Rubric Config`

`Category | Grade Type | Max Points | Scoring Type | Criterion Label | Criterion Text`

Max Points is populated on the first row of a category block only. GAS must forward-fill it when building the structured rubric object.

### `Group Grades`

`GroupKey | Period | Country | GroupWorkScore | GroupWorkComment | OutfitQualityScore | OutfitQualityComment | MetaphorScore | MetaphorComment | GroupStatus | LastSavedAt`

### `Individual Grades`

`Email | StudentName | Period | Country | EffortScore | EffortComment | ProfessionalismScore | ProfessionalismComment | ShowNightRole | ExtraCreditScore | ExtraCreditNote | GroupWorkOverride | GroupWorkOverrideNote | OutfitQualityOverride | OutfitQualityOverrideNote | MetaphorOverride | MetaphorOverrideNote | IndividualStatus | LastSavedAt | EmailSent | EmailSentAt`

### `Reflection Links`

`Email | StudentName | Period | Country | ReflectionLink | SubmittedAt`

### `Grade Rollup`

`Email | StudentName | Period | Country | GroupWorkFinal | OutfitQualityFinal | MetaphorFinal | GroupTotal | EffortScore | ProfessionalismScore | ExtraCreditScore | IndividualTotal | FinalScore | FinalPercent | EmailSent | EmailSentAt`
