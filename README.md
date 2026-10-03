# appscript_gmail_anfragen

## Deployment

The GitHub Actions deployment uses Clasp OAuth credentials. Run `clasp login`
with a Google account that has Editor access to the Apps Script project, then
add the complete contents of `~/.clasprc.json` as the repository Actions
secret `CLASP_CREDENTIALS_JSON`. This file contains a refresh token and must
not be committed to the repository.
