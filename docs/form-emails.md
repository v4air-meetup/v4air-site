# Form emails: where the texts live and how Typecase exports them

Reference for people and agents. The mailer is `sendMail` in `backend/Code.gs`; its rules are in the code comments.

## The contract (what the mailer reads)

| Item | Rule |
|---|---|
| Doc | Script property `EMAILS_DOC_ID`. Unset: "V4AIR form emails" (`1EkQNlVa_DGdI5lYP727lhxpAr0XhKTinMEud09Y_g9g`, the working Doc). petkout1 must be able to read it. |
| When read | Every run (every 5 minutes). An edit goes out from the next run. |
| One email | One tab. Tab name = email name, exactly (table below). Child tabs are read too. |
| Tab text | Optional lines first (e.g. the "Sent:" rule, ignored). Then one line `Subject: <subject>`. Everything after it is the body, as plain text. |
| Ignored tabs | Any tab without a `Subject:` line (e.g. "How it works"). |
| Placeholders | Single braces, filled per person: `{NAME}` `{ID}` `{LINK}` `{NOMINATE_LINK}` `{UNSUBSCRIBE_LINK}` `{SENT_DATE}` `{STARTED_DATE}`. Any other `{WORD}` blocks that email. |
| Required | Body has `{LINK}` and `{UNSUBSCRIBE_LINK}`. Subject not empty. |
| Broken text | That email is not sent; tab "Mail preview" (dry run) shows why. Doc unreadable: nothing is sent. |
| Late variant | Tab "<name> (late)", optional. Used when the email goes out more than 3 days after the form was sent or last saved. |

Tab names the mailer uses today:

| Tab | Sent to |
|---|---|
| Application confirmation | applicant, after sending the form |
| Application confirmation (late) | same, sent late (backlog) |
| Application resume | applicant with an unfinished form, 24 h after the last save |
| Application resume (late) | same, sent late (backlog) |
| Speaker confirmation | invited speaker, after sending the form |
| Speaker resume | invited speaker with an unfinished form, 24 h after the last save |

The other tabs of "V4AIR form emails" (reminders, nominations) are not sent yet.

## Typecase output for this Doc

| Setting | Value |
|---|---|
| Output format | `gdoc_tabs` (one Doc, one tab per document) |
| Document path rule | The exact tab name from the table above. No `/` (Typecase turns it into " · "), at most 50 characters. |
| Google folder | A folder in the V4AIR Drive tree (shared with petkout1). Typecase creates the Doc on the first export and updates it in place afterwards. |
| Text | First line `Subject: ...`, then a blank line, then the body. |
| Placeholders | Write the mailer's placeholders with single braces (`{NAME}`). Typecase only fills `{{fact}}`; single braces pass through unchanged. Never `{{NAME}}`. |
| Approval gate | `approved`: the export writes only reviewed text, and the export is what goes to applicants. |
| After the first export | Copy the new Doc's ID into script property `EMAILS_DOC_ID` (settings file, see `importSecrets` in `backend/Code.gs`). |

Rules:
- The exported Doc is export-only. Typecase replaces each tab's whole text on every export, so hand edits there are lost. Edit in Typecase.
- An email is not live until its tab exists in the Doc `EMAILS_DOC_ID` points to and its kind is switched on (`MAIL_KINDS`).
- Check a new export in tab "Mail preview" before it goes out: every email shows its filled-in text or the reason it is blocked.
