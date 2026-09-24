# Legacy mail templates

Drop the Outlook templates from `E:\O365AdminShared\EmailTemplates\` here, then run:

```powershell
pwsh scripts/Export-MailTemplates.ps1
```

That writes a `.html` (body) and `.txt` (subject, recipients, sender) next to each `.oft`. Those
are the source for the wording in `INotificationMailService` callers - currently
`SharedMailboxService.SendCreationConfirmationMailAsync` /
`SendOwnerChangeConfirmationMailAsync` / `SendRenameConfirmationMailAsync` and
`TeamsService.SendOwnerMailAsync`.

Templates that matter, per the legacy scripts:

| Template | Sent by | Ported |
|---|---|---|
| `SharedMailboxNew.oft` | `ShrMbxNew.ps1` | **yes** → `SharedMailboxService.SendCreationConfirmationMailAsync` |
| `SharedMailboxOwnershipChange.oft` | `ShrMbxChgOwner.ps1` | **yes** → `SharedMailboxService.SendOwnerChangeConfirmationMailAsync` |
| `SharedMailboxAlreadyExists.oft` | `ShrMbxNew.ps1` | not ported |
| `SharedMailboxRenamed.oft` | `RenameShrMbx.ps1` | **yes** → `SharedMailboxService.SendRenameConfirmationMailAsync` |
| `SharedMailboxRemoval.oft` | `ShrMbxRemove.ps1` | feature not ported |
| `DLNew.oft` | `NewDLGroup.ps1` | **yes** → `GroupNotificationMail.Creation` (+ `DynamicCreation`, no template) |
| `DLRenamed.oft` | `RenameDLGroup.ps1` | **yes** → `GroupNotificationMail.Renamed` |
| `DLRemoval.oft` | `RemoveDLGroup.ps1` | **yes** → `GroupNotificationMail.Removal` |
| `DistributionListOwnershipChange.oft` | `UpdateGroupOwnership.ps1` | **yes** → `GroupNotificationMail.OwnershipChanged` |
| `DLMembershipReplace.oft` | `ReplGrpMembers.ps1` | **yes** → `GroupNotificationMail.MembershipReplaced` |
| `DLRestrictedAccessChange.oft` | `ModifyDLAuthUsers.ps1` | **yes** → `GroupNotificationMail.AuthorizedSendersChanged` |
| `DLRestrictedAccessGranted.oft` | `DistributionSecurityGroupMenu.ps1` | not used - see below |
| `RoomorResourceNewSite.oft` | `RoomResourceNewForm.ps1` | **yes** → `RoomNotificationMail.Creation`, new room list + general use |
| `RestrictedRoomsNewSite.oft` | `RoomResourceNewForm.ps1` | **yes** → same, new room list + restricted |
| `RoomorResourceAdditions.oft` | `RoomResourceNewForm.ps1` | **yes** → same, existing list + general use |
| `RestrictedRoomAdditions.oft` | `RoomResourceNewForm.ps1` | **yes** → same, existing list + restricted |
| `RoomResourceRemoval.oft` | `RoomResourceRemove.ps1` | **yes** → `RoomNotificationMail.Removal` |
| `RROOPChanges.oft` | `RoomResourceOOPChanges.ps1` | **yes** → `RoomNotificationMail.UsersAuthorized` |
| `RoomorResourceRename.oft` | `RoomResourceRename.ps1` | feature not ported |

The `.oft` files themselves don't need to be committed - the exported `.txt` is what the code is
derived from.

## Deviations from the originals

- **The embedded screenshots are not carried over.** `SharedMailboxNew.oft`'s showed the Outlook
  2010 address book with the sending admin's own mail address visible in it. The surrounding text
  explains the step without it.
- **The ServiceNow links are settings, not code** (*Settings → Shared mailboxes → Confirmation
  e-mail links*). The two templates linked the same three articles under two different URL formats
  (`?id=kb_article&sysparm_article=KB00…` vs `?sys_kb_id=…&id=kb_article_view`), so one set is
  stale. The defaults use the first form - verify them.
- **Placeholders are filled from the request** rather than by hand: the group list only names the
  tiers actually created, and `MBX.???.ED` becomes the real group name.
- The retention sentence only appears when a retention policy is configured, and the
  external-senders sentence follows the checkbox on the create form.
- **The room templates' blank line becomes the room list name.** Each of the four creation
  templates had an empty indented line where the operator pasted a Room Finder screenshot; the
  room list is what someone actually needs in order to find the rooms. Equipment has no room list,
  so the block is dropped rather than naming something that doesn't exist.
- **The room booking numbers come from the applied policy**, not the hardcoded "90 days / 24 hr"
  the originals carried - a Hoteling room would otherwise be described wrongly.
- **Room mails go to a new "Requester" field.** Rooms have no owner the way a shared mailbox does,
  and the originals were addressed by hand (every `To:` is empty). Leaving it blank sends no mail.
- **`RROOPChanges.oft` is not about booking policies** despite the name - its subject is
  "Additional Users Authorized to Reserve …", so it maps to a membership change, and only to
  *adding* users to a restricted room's authorized-users group. Removals and delegate-group changes
  have no template and send nothing.
- **A details change has no legacy template**; that wording is new and lists what actually changed.
- Obvious grammar errors in the originals were fixed ("The existing room delegates have designated
  to manage…" → "have been designated to manage…", "you have been access to" → "have been given
  access to").
- **`DistributionListOwnershipChange.oft` was a "Select one:" template**, holding four alternative
  paragraphs the operator chose between plus the editorial note "Keep this ... if this is not a
  unified group". `OwnershipChanged` picks by `ListChangeMode` instead, and drops the membership
  link for Microsoft 365 groups - which is what that note asked for.
- **`DLRestrictedAccessGranted.oft` is not used.** It is the same message as
  `DLRestrictedAccessChange.oft` but addressed to the people who were granted access ("*you* have
  been given access") rather than to the requester. Everything here goes to the requester. If those
  individuals should be notified directly as well, that is a separate recipient decision.
- **`SharedMailboxRenamed.oft` lists only the groups the mailbox actually has.** The original
  always named all three tiers (and misspelled two of them as "MBX. NewNameOfMailbox.AU"), so a
  mailbox with an .ED group only got a mail pointing at two groups that do not exist. It also gains
  a sentence saying the previous address is kept as an alias, which is what the rename now does.
- **Group alias changes send nothing** - no legacy template covered them.

