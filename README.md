# Exchange-Online-Migration-Toolkit (PowerShell)

Please stop by https://m365adminstools.com for more info and IT engineering tools.

An interactive menu for the repetitive per-user tasks in a hybrid Exchange to Exchange Online migration: stamping the on-premises AD attributes that route mail to the cloud mailbox, scheduling move request completion, updating UPNs to the vanity domain, checking licence assignment, and checking move status. Each task runs against a single user or against a list of accounts in bulk.

Every action is written to a log file with a timestamp.

All the environment-specific values live in one configuration block at the top. Update that block once per engagement and the rest of the script follows.

<!-- Add a screenshot of the menu here, then uncomment:
![Menu](docs/images/menu.png)
-->

## This script makes changes

Unlike the reporting tools in this collection, this one writes. It modifies Active Directory user attributes, adds group memberships, and schedules mailbox move completion in Exchange Online.

Three specific points to understand before running it:

- **There is no preview or confirmation.** The bulk options iterate the entire account list and apply every change without prompting. There is no `-WhatIf`.
- **Move completion is scheduled with `-AcceptLargeDataLoss` and a bad item limit of 1000.** This is a deliberate choice for migrations where a small number of corrupt items must not block a cutover, but it means up to 1000 items per mailbox can be skipped without stopping the move. Lower `BadItemLimit` in the configuration block if that is not what you want.
- **Completion time is calculated from today's date.** If the configured completion time has already passed when you schedule the batch, the move completes immediately rather than at the next occurrence of that time. Schedule before the configured time, or set the time to a value still ahead of you.

Test against a pilot batch before running it across a migration.

## Requirements

| Item | Requirement |
|---|---|
| PowerShell | 7, elevated |
| Modules | `ExchangeOnlineManagement`, `ActiveDirectory`, `Microsoft.Graph.Users` |
| Machine | Domain joined, or with RSAT AD tools and a route to a domain controller |
| Rights | Permission to modify the target AD users and to add members to the licence group. In the tenant, a role that can run move requests and read user licence assignment |
| Environment | Hybrid Exchange with directory synchronization, where on-premises AD is the source of authority |

The script checks for the three modules at startup and warns without stopping, so a missing module surfaces as a failed menu option rather than a refusal to start.

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
```

## Configuration

Edit the `$Config` block at the top of the script before first use.

| Key | Example | Purpose |
|---|---|---|
| `AdminUPN` | `admin@contoso.com` | Account used to connect to Exchange Online, IPPS, and Graph |
| `PrimaryDomain` | `contoso.com` | Vanity domain used to build each user's UPN |
| `RoutingDomain` | `contoso.mail.onmicrosoft.com` | Tenant routing domain used to build targetAddress and the routing proxy address |
| `LicenseGroup` | `O365_E3_License` | AD security group that grants the target licence |
| `AccountListFile` | `C:\Temp\Account_List.txt` | Text file of sAMAccountNames, one per line |
| `LogFile` | `C:\Temp\Migration_Output.txt` | Log for every action and error |
| `CompletionTime` | `05:00 AM` | Local time used to build the `CompleteAfter` value, converted to UTC |
| `BadItemLimit` | `1000` | Bad item limit applied when scheduling completion |

The account list is plain text, one sAMAccountName per line. Blank lines are ignored.

```
jsmith
mjones
rpatel
```

## Quick start

```powershell
# Edit the $Config block first, then:
.\Migration_Process.ps1
```

Then, in the menu:

1. Choose option 1 to connect. This opens sign-in prompts for Exchange Online, the Security and Compliance endpoint, and Microsoft Graph.
2. Run the option you need, single or bulk.
3. The current tenant domain, routing domain, account list path, and log path are shown at the bottom of the menu, so you can confirm you are pointed at the right environment before acting.

If the script is blocked on first run:

```powershell
Unblock-File .\Migration_Process.ps1
```

## Menu reference

| Option | Action | What it changes |
|---|---|---|
| 1 | Connect to Exchange Online, IPPS, and Graph | Nothing |
| 2 | Bulk, complete migrated mailboxes in AD | Per user: sets UPN to the primary domain, adds an `smtp:` proxy address on the routing domain, and replaces `targetAddress` with the routing address |
| 3 | Single, complete migrated mailbox in AD | The same as option 2 for one user, and additionally adds the user to the licence group |
| 4 | Bulk, schedule synced mailbox completion | Sets `CompleteAfter` on each move request, with the configured bad item limit and `-AcceptLargeDataLoss` |
| 5 | Single, schedule synced mailbox completion | The same for one user |
| 6 | Bulk, update UPN to primary domain | Sets the UPN only |
| 7 | Single, update UPN to primary domain | Sets the UPN only |
| 8 | Single, check licence assignment | Nothing. Reads display name and assigned licence SKUs from Graph |
| 9 | Bulk, check licence assignments | Nothing. Reads licence counts for every account in the list |
| 0 | Bulk, check migration status | Nothing. Reads move request name and status |
| Z | Single, check migration status | Nothing |
| Q | Quit | Nothing |

Options 8, 9, 0, and Z are read-only and safe to run at any time.

## Note on option 3 versus option 2

The single-user option adds the user to the licence group. The bulk option does not. This is intentional in a workflow where licensing is handled separately for a batch, but it means the two options are not equivalent. If you expect group membership to be applied in bulk, add it, or run the group membership step separately.

## Logging

Every action, success and failure, is appended to the log file with a timestamp, including the values applied per user. Errors are caught per user, so one failure does not halt a bulk run. Review the log after every batch rather than relying on the console, because a long bulk run scrolls past.

## Troubleshooting

**`Module '<name>' is not installed.`**

Shown at startup for each missing module. Install it with the commands in the Requirements section and restart the script.

**`Account list not found at '<path>'. Create it first.`**

The file named in `AccountListFile` does not exist. Create it, one sAMAccountName per line.

**`The specified value already exists` when adding a proxy address**

The routing proxy address is already on the account, usually because option 2 or 3 was run twice for that user. The error is caught and logged, the other attribute changes still apply, and no harm is done.

**`The operation couldn't be performed because object '<user>' couldn't be found`**

Either the sAMAccountName is wrong, or there is no move request for that user in the tenant. Check the spelling in the account list, and confirm the migration batch exists.

**A move completes immediately instead of at the scheduled time**

The configured completion time had already passed for today when the batch was scheduled, so the `CompleteAfter` value was in the past. See the note at the top of this README.

**Licence check returns nothing for a user**

`Get-MgUser` is queried by the UPN built from the sAMAccountName and the primary domain. If the user's actual UPN differs from that pattern, the lookup fails. Check the UPN in the tenant.

**Connect fails for Microsoft Graph**

The script requests the `User.Read.All` scope. If your tenant requires admin consent for that scope and it has not been granted, sign-in fails. Have a Global Administrator consent to the permission once.

## Limitations

- Built for hybrid Exchange with directory synchronization and on-premises AD as the source of authority. It is not suitable for a cloud-only tenant, a cross-tenant migration, or a cutover migration without on-premises AD.
- User identifiers are assumed to follow `sAMAccountName@PrimaryDomain`. Environments where the UPN does not follow that pattern need the string building adjusted.
- Only one primary domain and one routing domain per run. A multi-domain migration needs a separate configuration block per domain.
- The script does not create move requests, create migration batches, or perform the mailbox move. It manages the surrounding per-user tasks.
- No rollback. Record the previous UPN and targetAddress values before running if you need to be able to revert.

## Related

- Free Microsoft 365, Active Directory, and Veeam tools at [m365admintools.com](https://m365admintools.com)

## Author

Charles Arconi, [m365admintools.com](https://m365admintools.com)

## License

MIT. See [LICENSE](LICENSE).

