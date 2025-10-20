/**
 * Configuration variables. Adjust these to match your needs.
 */
const CONFIG = {
  TARGET_SUBJECT: "New Purchase Order",// The exact subject line to match

  SHARED_DRIVE_ID: "",
  FolderID: "", // must be created ahead of time
   
  PROCESSED_LABEL_NAME: "attachments-saved-to-drive" // The label name to apply to processed emails to avoid duplication.
};

// ----------------------------------------------------------------------
// Main Function
// ----------------------------------------------------------------------

/**
 * Searches for sent emails matching the subject, saves attachments to Drive,
 * and marks the threads with a label.
 */
function saveAttachmentsToDrive() {
  // 1. Get or create the target folder in the Shared Drive.
  // const folderId = getOrCreateSharedDriveFolder(CONFIG.SHARED_DRIVE_ID, CONFIG.DRIVE_FOLDER_NAME); //doesnt work as intended
  const folderId = CONFIG.FolderID

  // 2. Get or create the label for processed emails.
  const label = getOrCreateGmailLabel(CONFIG.PROCESSED_LABEL_NAME);

  // 3. Construct the Gmail search query.
  // Query excludes threads that have already been processed and only looks at sent mail.
  const searchQuery = `from:me subject:"${CONFIG.TARGET_SUBJECT}" -label:${CONFIG.PROCESSED_LABEL_NAME}`;

  // Fetch threads matching the query.
  const threads = GmailApp.search(searchQuery);

  Logger.log(`Found ${threads.length} threads matching the query: ${searchQuery}`);

  threads.forEach(thread => {
    // SECURITY CHECK: Since the search query *should* filter these out, 
    // we use a check at the thread level to confirm we are not reprocessing.
    // getLabels() is the correct function on the GmailThread object.
    const threadLabels = thread.getLabels();
    const isThreadProcessed = threadLabels.some(threadLabel => threadLabel.getName() === CONFIG.PROCESSED_LABEL_NAME);

    if (isThreadProcessed) {
      Logger.log(`Skipping thread ID: ${thread.getId()}. Already processed by search criteria.`);
      return;
    }

    let attachmentsSavedInThread = false;

    // Process messages within the thread (usually just one message for sent mail).
    const messages = thread.getMessages();

    messages.forEach(message => {
      // Basic check: Skip if trashed or draft.
      if (message.isInTrash() || message.isDraft()) {
        return;
      }

      const attachments = message.getAttachments();

      if (attachments.length > 0) {
        attachments.forEach(attachment => {
          if (attachment.getContentType() === 'application/pdf') {
            try {
              // *** NEW LOGIC: Use Advanced Drive Service to save file to Shared Drive ***
              const fileMetadata = {
                name: attachment.getName(),
                parents: [folderId], // Set the folder ID as the parent
              };

              // Insert the file using the Advanced Service, enabling Shared Drive support
              Drive.Files.create(fileMetadata, attachment, {
                supportsAllDrives: true
              });

              Logger.log(`Saved PDF attachment: ${attachment.getName()} from message ID: ${message.getId()}`);
              attachmentsSavedInThread = true;
            } catch (e) {
              Logger.log(`ERROR: Could not save PDF attachment ${attachment.getName()}. Error: ${e.toString()}`);
            }
          } else {
            Logger.log(`Skipped attachment: ${attachment.getName()}. Content type is ${attachment.getContentType()}, not PDF.`);
            attachmentsSavedInThread = true; // flag anyway so it doesnt rerun
          }
        });
      }
    });

    // Mark the entire THREAD as processed ONLY if attachments were successfully saved.
    // Applying the label to the thread is the most reliable way to prevent reprocessing 
    // by the next run of the GmailApp.search() function.
    if (attachmentsSavedInThread) {
      thread.addLabel(label);
      Logger.log(`Applied label to thread ID: ${thread.getId()}`);
    }
  });
}

// ----------------------------------------------------------------------
// Helper Functions (Unchanged)
// ----------------------------------------------------------------------

/**
 * Helper function to find an existing folder within a Shared Drive or create a new one.
 * Uses the Drive Advanced Service (Drive API).
 * @param {string} driveId The ID of the Shared Drive.
 * @param {string} folderName The name of the folder to find/create.
 * @returns {string} The ID of the found or created folder.
 */
// function getOrCreateSharedDriveFolder(driveId, folderName) { // NOT WORKING, always creates a new folder as it doesnt find the exising one
//   // 1. Search for the folder by name within the Shared Drive ID
//   // REVISED QUERY: We are simplifying the query and relying on the `driveId` parameter 
//   // to perform the search within the Shared Drive.
//   const searchResults = Drive.Files.list({
//     q: `name = '${folderName}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
//     corpora: 'drive', // Indicates searching within a specific drive (Shared Drive)
//     driveId: driveId, // Specifies the ID of the Shared Drive
//     includeItemsFromAllDrives: true,
//     supportsAllDrives: true,
//     maxResults: 1
//   });

//   if (searchResults.items && searchResults.items.length > 0) {
//     Logger.log(`Found existing folder: ${folderName}`);
//     return searchResults.items[0].id; // Return the existing folder ID
//   } else {
//     // 2. Create the folder if it doesn't exist
//     const folderMetadata = {
//       name: folderName,
//       mimeType: 'application/vnd.google-apps.folder',
//       parents: [driveId], // Set the Shared Drive ID as the parent
//     };

//     // Ensure supportsAllDrives is true when inserting into a Shared Drive
//     const newFolder = Drive.Files.create(folderMetadata, null, {
//           corpora: 'drive', // Indicates searching within a specific drive (Shared Drive)
//     driveId: driveId, // Specifies the ID of the Shared Drive
//       supportsAllDrives: true
//     });
//     Logger.log(`Created new folder: ${folderName}`);
//     return newFolder.id; // Return the new folder ID
//   }
// }
/**
 * Helper function to find an existing Gmail label or create a new one.
 * @param {string} labelName The name of the label.
 * @returns {GoogleAppsScript.Gmail.GmailLabel} The Gmail label object.
 */
function getOrCreateGmailLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}
