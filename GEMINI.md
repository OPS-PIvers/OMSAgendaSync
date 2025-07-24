# Deployment Process for OMSAgendaSync

This document outlines the standard deployment process for the OMSAgendaSync web application. Following these steps ensures that changes are properly versioned in Git and deployed to Google Apps Script (GAS).

## Process Steps

After *every* update to the codebase (e.g., `Code.js`, `index.html`, `appsscript.json`), the following deployment steps *must* be performed:

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

3.  **Push to Remote Repository:**
    Push your local commits to the remote Git repository (e.g., GitHub).
    ```bash
    git push
    ```

4.  **Push to Google Apps Script (GAS) Project:**
    I will automatically perform `clasp push` after a successful `git push`.
    ```bash
    clasp push
    ```

5.  **Update Active Web App Deployment:**
    I will automatically perform `clasp redeploy` after a successful `clasp push`.
    ```bash
    clasp redeploy AKfycbzWZD2iUIPMwpJAJ5fE53_372YP_sz4XR2U6nYl0dQjsImIcSf_8F_-qzEn7rS3tVWzdA --versionNumber <LATEST_VERSION_NUMBER>
    ```
    (Replace `<LATEST_VERSION_NUMBER>` with the actual latest version number. I will determine this automatically.)

    **Important:** After updating the deployment, it's often necessary to clear your browser's cache or open the web app in an incognito/private window to ensure you are viewing the latest deployed version.
