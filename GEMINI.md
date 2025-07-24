# Deployment Process for OMSAgendaSync

This document outlines the standard deployment process for the OMSAgendaSync web application. Following these steps ensures that changes are properly versioned in Git and deployed to Google Apps Script (GAS).

## Process Steps

After *every* update to the codebase (e.g., `Code.js`, `index.html`, `appsscript.json`), the following deployment steps *must* be performed *by Gemini*:

1.  **Stage Changes:**
    Add all modified files to the Git staging area.
    ```bash
    git add .
    ```

2.  **Commit Changes:**
    Commit the staged changes with a descriptive commit message. Follow conventional commit guidelines (e.g., `FEAT:`, `FIX:`, `CHORE:`).
    ```bash
    git commit -m "FEAT: Your descriptive commit message here"
    ```

3.  **Push to Remote Repository & Deploy to Google Apps Script:**
    Push your local commits to the remote Git repository (e.g., GitHub). After a successful `git push`, I will automatically perform the `clasp push` and `clasp redeploy` steps to update your Google Apps Script project and web app deployment.
    ```bash
    git push
    ```
    **Note:** I will manage the `CLASP_DEPLOYMENT_ID` and versioning internally. You do not need to manually run `clasp push` or `clasp redeploy` after I perform a `git push`.

    **Important:** After updating the deployment, it's often necessary to clear your browser's cache or open the web app in an incognito/private window to ensure you are viewing the latest deployed version.
