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
| Placeholders | Single braces, filled per person: `{NAME}` `{ID}` `{LINK}` `{NOMINATE_LINK}` `{UNSUBSCRIBE_LINK}` `{SENT_DATE}` `{STARTED_DATE}` `{NOMINATOR}` `{ACCEPT_LINK}`. Any other `{WORD}` blocks that email. |
| `{LINK}` | Applications and speakers: the personal link to their form. Nomination confirmation and nominate reminders: the nominate page. Nominee emails: the apply page. |
| Required | Body has `{UNSUBSCRIBE_LINK}` and `{LINK}` (nominee emails: `{ACCEPT_LINK}` instead). Subject not empty. |
| Broken text | That email is not sent; tab "Mail preview" (dry run) shows why. Doc unreadable: nothing is sent. |
| Late variant | Tab "<name> (late)", optional. Used when the email goes out more than 3 days after the form was sent or last saved. |

Tab names the mailer uses, in priority order (one email per person per day: the first one due goes). `kind` is the value for `MAIL_KINDS`.

| Tab | Kind | Sent to | When |
|---|---|---|---|
| Application confirmation (+ late) | confirm | applicant | form sent |
| Speaker confirmation | confirm | invited speaker | form sent |
| Nomination confirmation | confirm | the nominator | nomination sent |
| Deadline reminder: 1 week | deadline1 | applicant with an unfinished form | 8 Jan 2027 from 09:00 |
| Deadline reminder: last day | deadline2 | same | 14 Jan 2027 from 09:00 |
| Application resume (+ late) | resume | same | 1 day after the last save |
| Application resume 2 | resume2 | same | 7 days after the last save, after resume; not from 5 Jan 2027 |
| Speaker resume | resume | speaker with an unfinished form | 1 day after the last save |
| Speaker resume 2 | resume2 | same | 4 days after the last save, after resume |
| Speaker resume 3 | resume3 | same | 10 days after the last save, after resume 2 |
| Nominee invitation | invite | nominee with an email address, not declined, not applied | right after the nomination |
| Nominee reminder | remind | same | 7 days after the invitation |
| Nominate reminder 1 | nominate1 | applicant who nominated nobody | 3 days after applying, after the confirmation |
| Nominate reminder 2 | nominate2 | same | 14 days after applying |
| Nominate reminder 3 | nominate3 | same | 28 days after applying |

For every email: not to unsubscribed addresses (confirmations excepted) or test addresses; nothing but confirmations after
15 January 2027; each reminder once per address; rows saved before go-live count as saved at go-live. Nominee details are
deleted after 30 days by the daily check (`purgeNominees`), in live mode or with `NOMINEE_PURGE=true`.

Test mode (live mail to test addresses only): `MAIL_MODE=live` plus `MAIL_ONLY` (comma separated, `*` matches any text,
e.g. `chlupp+*@natur.cuni.cz`) sends real emails with working links to those addresses only; every other row stays untouched
and `MAIL_START` is not stored. `MAIL_TEST_DAY_MINUTES=10` makes a "day" 10 minutes for those addresses (reminders, late texts,
one email a day, the 30-day deletion). Delete the test rows afterwards, then remove both settings before the real go-live.

Reading every email before go-live: `python3 tests/mail/fixture-from-doc.py`, then `node tests/mail/samples.mjs --out samples.txt`
(all 17 with made-up data), or `--send <address>` to receive them through Resend.

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
