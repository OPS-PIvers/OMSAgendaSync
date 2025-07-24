# Deployment Process for OMSAgendaSync

This document outlines the standard deployment process for the OMSAgendaSync web application. Following these steps ensures that changes are properly versioned in Git and deployed to Google Apps Script (GAS).

## Process Steps

After making any changes to the codebase (e.g., `Code.js`, `index.html`, `appsscript.json`):

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
    Use `clasp` to push your local code to the associated Google Apps Script project. This updates the script project's files.
    ```bash
    clasp push
    ```

5.  **Update Active Web App Deployment:**
    After `clasp push`, you need to update the deployed web app to reflect the latest code. This is crucial because `clasp push` only updates the script project, not the deployed version.

    To update the deployment, you will use the `clasp deploy` command with the deployment ID.

    **Setting the Deployment ID:**
    For security and convenience, it is highly recommended to export your web app's deployment ID as an environment variable in your terminal session or shell configuration (e.g., `.bashrc`, `.zshrc`).

    **Action:**
    Manually set the `CLASP_DEPLOYMENT_ID` environment variable with your deployment ID:
    ```bash
    export CLASP_DEPLOYMENT_ID="AKfycbzWZD2iUIPMwpJAJ5fE53_372YP_sz4XR2U6nYl0dQjsImIcSf_8F_-qzEn7rS3tVWzdA"
    ```

    Then, update the deployment using the environment variable:
    ```bash
    clasp deploy --deploymentId $CLASP_DEPLOYMENT_ID --versionNumber $(clasp versions | tail -n 1 | awk '{print $1}')
    ```
    This command creates a new version and updates the specified deployment to use that new version. The `$(clasp versions | tail -n 1 | awk '{print $1}')` part automatically fetches the latest version number from your GAS project.

    **Important:** After updating the deployment, it's often necessary to clear your browser's cache or open the web app in an incognito/private window to ensure you are viewing the latest deployed version.
