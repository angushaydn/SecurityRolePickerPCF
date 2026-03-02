# SecurityRolePicker PCF

SecurityRolePicker is a React-based Power Apps Component Framework (PCF) control for managing Dataverse security roles on `team` and `systemuser` records.

It provides:
- A searchable **Available Roles** list.
- A searchable **Assigned Roles** list.
- Role **associate** (`>`) and **disassociate** (`<`) actions.
- A live **privilege matrix** showing effective rights by table/entity.

## What The Control Does

When hosted on a team or user form, the control:
1. Detects the current record ID and type (`team` or `systemuser`).
2. Retrieves current principal details and currently assigned roles.
3. Retrieves the full list of Dataverse roles.
4. Calculates and renders a privilege matrix from Dataverse privilege data.
5. Lets the user add/remove roles and immediately refreshes the matrix.

## UI Behavior

The UI has two main panels:
- Left panel: role management
- Right panel: privilege matrix

Role management panel:
- Search box filters both lists by role name.
- Selecting a role in **Available Roles** enables `>`.
- Selecting a role in **Assigned Roles** enables `<`.
- After an add/remove action, list selections are updated and matrix is reloaded.

Privilege matrix panel:
- Rows = entities (derived from privilege names).
- Columns = `Create`, `Read`, `Write`, `Delete`, `Append`, `Append To`, `Assign`, `Share`.
- Each cell shows a colored dot representing the **highest privilege depth** for that right on that entity.

Depth color mapping:
- `basic` -> gray `#797775`
- `local` -> blue `#2266E3`
- `deep` -> brown/red `#A52A2A`
- `global` -> green `#107C10`
- no privilege -> white `#FFFFFF`

## Dataverse Interactions

### Principal and role retrieval

On load, the control uses `context.webAPI` to retrieve:
- Principal record (`team` or `systemuser`) with expanded role relationship.
- All roles from `role` table.

Relationships:
- Team roles: `teamroles_association`
- User roles: `systemuserroles_association`

### Privilege retrieval

To build the matrix, the control executes one of:
- `RetrieveTeamPrivileges` (for teams)
- `RetrieveUserPrivileges` (for users)

Then it parses privilege names in `prv<Verb><Entity>` format to map:
- Verb -> matrix column
- Entity suffix -> matrix row

Supported verbs:
- `Create`
- `Read`
- `Write`
- `Delete`
- `Assign`
- `AppendTo`
- `Append`
- `Share`

If multiple privileges exist for the same entity/right, the highest depth wins.

### Role assignment actions

Add role (`>`):
- Executes Dataverse `Associate` request with principal/role relationship.

Remove role (`<`):
- Executes Dataverse `Disassociate` request with principal/role relationship.

After either action:
- Local state is updated.
- Matrix is reloaded from Dataverse for fresh effective privileges.

## Error Handling

Errors are surfaced through `Xrm.App.addGlobalNotification` when available.
If not available, the message is logged to console.

Error messages are normalized from:
- Standard `Error.message`
- Dataverse `raw` payloads containing nested `message`

## Technical Architecture

Main files:
- `SecurityAccessPicker/index.ts`
- `SecurityAccessPicker/BentoGrid.tsx`
- `SecurityAccessPicker/BentoGrid.css`
- `SecurityAccessPicker/ControlManifest.Input.xml`

Architecture notes:
- PCF control type is `virtual`.
- React rendering is wrapped with Fluent v9 providers.
- Container height is clamped to allocated host height and viewport availability.
- CSS grid layout supports nested scroll regions and sticky headers.
- Mobile layout collapses to one column at `max-width: 900px`.

## Manifest and Runtime Requirements

Manifest (`ControlManifest.Input.xml`) includes:
- Namespace: `Am.Security`
- Constructor: `SecurityAccessPicker`
- Required features: `Utility`, `WebAPI`
- Platform libraries: React `16.14.0`, Fluent `9.4.0`

Note:
- The manifest currently includes a required placeholder bound property `sampleProperty`. The control logic does not use this property directly.

### Import Solution

1. Download the solution file:
   - **Unmanaged**: `Solutions/SecurityRolePickerPCF_1_0_0_1.zip`
   

2. Go to [Power Apps](https://make.powerapps.com)

3. Navigate to **Solutions** > **Import solution**

4. Select the downloaded `.zip` file and follow the import wizard

## Deploying To Dataverse

Typical flow:
1. Build the control.
2. Add/update in a Dataverse solution.
3. Push/import to target environment.
4. Place on `team` or `systemuser` forms.
5. (Recommended) Set on a brand new tab on form with "Expand first component to full tab" for the best view/management

For publisher ownership in Dataverse:
- Publisher is determined by the target solution/component in the environment, not just by CLI prefix arguments.
- Use the correct solution unique name if you require a specific publisher.

## Known Assumptions and Limits

- Control expects to run in context of a `team` or `systemuser` record.
- Principal detection falls back to page context and defaults entity type to `team` if type cannot be resolved.
- Privilege matrix depends on Dataverse privilege naming convention (`prv<Verb><Entity>`).
- Display labels for entities are prettified from logical suffixes and may not match localized display names.

## Repository Structure

- `SecurityAccessPicker/` - control source (TSX/CSS/manifest)
- `out/` - build output
- `obj/` - intermediate build artifacts
- `Solutions/` - solution packaging artifacts
- `SecurityRolePicker.pcfproj` - MSBuild project for PCF packaging

# Preview
<img width="2278" height="1165" alt="image" src="https://github.com/user-attachments/assets/7d713612-9432-42ed-82c4-0e3357578f84" />

