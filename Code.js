/**
 * @OnlyCurrentDoc
 * A script to merge multiple presentations from a Google Drive folder.
 * VERSION 9.1: Added sorting by upload/creation date.
 */

// --- CONFIGURATION ---
const FOLDER_ID = '1jFHtyulML-HM9MxPHp69oZhS2p2G7XaS';
const MERGED_FILE_NAME_PREFIX = 'Combined Final Presentation';
// ---------------------

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

// ====================================================================================
// == PUBLIC FUNCTIONS - CHOOSE THE MERGE ORDER YOU WANT ==
// ====================================================================================

/** Merges files using a "natural sort" on their names in ASCENDING order (e.g., Topic 1, 2, 10). */
function mergeByName_Ascending() {
  const files = _getAndSortFiles('name', false);
  _startMergeProcess(files, ' (By Name - Ascending)');
}

/** Merges files using a "natural sort" on their names in DESCENDING order (e.g., Topic 10, 2, 1). */
function mergeByName_Descending() {
  const files = _getAndSortFiles('name', true);
  _startMergeProcess(files, ' (By Name - Descending)');
}

/** Merges files based on their modification date in DESCENDING order (Newest first). */
function mergeByModifiedDate_Descending() {
  const files = _getAndSortFiles('modified', true);
  _startMergeProcess(files, ' (By Modified Date - Descending)');
}

/** Merges files based on their UPLOAD date in ASCENDING order (Oldest upload first). */
function mergeByUploadDate_Ascending() {
  const files = _getAndSortFiles('created', false);
  _startMergeProcess(files, ' (By Upload Date - Ascending)');
}

/** Merges files based on their UPLOAD date in DESCENDING order (Newest upload first). */
function mergeByUploadDate_Descending() {
  const files = _getAndSortFiles('created', true);
  _startMergeProcess(files, ' (By Upload Date - Descending)');
}

/** CANCELS any merge process that is currently in progress. */
function cancelMergeProcess() {
  _clearPropertiesAndTriggers();
  Logger.log('CANCELLED: All background processes have been stopped and cleaned up.');
}


// ====================================================================================
// == INTERNAL FUNCTIONS - DO NOT RUN THESE DIRECTLY ==
// ====================================================================================

function _startMergeProcess(filesToProcess, suffix) {
  _clearPropertiesAndTriggers();
  if (!filesToProcess || filesToProcess.length === 0) return;
  
  const fileIdsToProcess = filesToProcess.map(file => file.getId());
  const outputFileName = `${MERGED_FILE_NAME_PREFIX}${suffix}`;
  
  if (DriveApp.getFilesByName(outputFileName).hasNext()) {
    Logger.log(`A file named "${outputFileName}" already exists. Please delete or rename it and start again.`);
    return;
  }
  
  const mergedPresentation = SlidesApp.create(outputFileName);
  SCRIPT_PROPERTIES.setProperties({
    'destinationId': mergedPresentation.getId(),
    'filesToProcess': JSON.stringify(fileIdsToProcess),
    'runCount': '1'
  });

  ScriptApp.newTrigger('_continueMergeProcess').timeBased().after(60 * 1000).create();
  Logger.log(`SUCCESS: Started new merge process. Output will be "${outputFileName}". The script will now run in the background.`);
}

/**
 * Universal sorting function.
 * @private
 */
function _getAndSortFiles(sortBy = 'name', descending = false) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const filesArray = [];
    const filesIterator = folder.getFiles();
    while (filesIterator.hasNext()) {
      const file = filesIterator.next();
      if (file.getName().endsWith('.pptx') || file.getMimeType() === MimeType.GOOGLE_SLIDES) {
        filesArray.push(file);
      }
    }
    if (filesArray.length === 0) { Logger.log('No compatible files found.'); return null; }

    filesArray.sort((a, b) => {
      let result;
      switch (sortBy) {
        // --- NEW: Sort by upload/creation date ---
        case 'created':
          result = a.getDateCreated() - b.getDateCreated();
          break;
        // ------------------------------------------
        case 'modified':
          result = a.getLastUpdated() - b.getLastUpdated();
          break;
        case 'size':
          result = a.getSize() - b.getSize();
          break;
        case 'name':
        default:
          const re = /(\d+)|(\D+)/g;
          const chunksA = a.getName().match(re) || [];
          const chunksB = b.getName().match(re) || [];
          const len = Math.min(chunksA.length, chunksB.length);
          for (let i = 0; i < len; i++) {
            const chunkA = chunksA[i];
            const chunkB = chunksB[i];
            const numA = parseInt(chunkA, 10);
            const numB = parseInt(chunkB, 10);
            if (!isNaN(numA) && !isNaN(numB)) {
              if (numA !== numB) { result = numA - numB; break; }
            } else if (chunkA !== chunkB) {
              result = chunkA.localeCompare(chunkB); break;
            }
          }
          if (result === undefined) { result = chunksA.length - chunksB.length; }
          break;
      }
      return descending ? -result : result;
    });
    return filesArray;
  } catch (e) {
    Logger.log(`Error getting or sorting files: ${e.toString()}`);
    return null;
  }
}

// The _continueMergeProcess, _clearPropertiesAndTriggers, and appendSlideWithRetry functions
// are identical to the previous version and are included here for completeness.

function _continueMergeProcess() {
  const startTime = new Date();
  const properties = SCRIPT_PROPERTIES.getProperties();
  const destinationId = properties.destinationId;
  const filesToProcess = JSON.parse(properties.filesToProcess || '[]');
  
  if (!destinationId || filesToProcess.length === 0) { _clearPropertiesAndTriggers(); return; }

  Logger.log(`CONTINUING MERGE (Run #${properties.runCount}): ${filesToProcess.length} files remaining.`);
  const mergedPresentation = SlidesApp.openById(destinationId);

  while (filesToProcess.length > 0 && (new Date() - startTime) < (1000 * 60 * 4.5)) {
    const fileId = filesToProcess.shift();
    const file = DriveApp.getFileById(fileId);
    
    Logger.log(`Processing file: "${file.getName()}"`);
    let sourcePresentationId = null;
    let isTempFile = false;
    try {
      if (file.getName().endsWith('.pptx')) {
        const tempSlidesFile = Drive.Files.copy({ title: `[TEMP] ${file.getName()}`, mimeType: MimeType.GOOGLE_SLIDES }, file.getId());
        sourcePresentationId = tempSlidesFile.id; isTempFile = true;
      } else {
        sourcePresentationId = file.getId();
      }
      const sourcePresentation = SlidesApp.openById(sourcePresentationId);
      const sourceSlides = sourcePresentation.getSlides();
      for (let i = 0; i < sourceSlides.length; i++) {
        appendSlideWithRetry(mergedPresentation, sourceSlides[i], i + 1, file.getName());
        Utilities.sleep(250);
      }
    } catch (e) {
      Logger.log(`!! Could not process file "${file.getName()}". Skipping. Error: ${e}`);
    } finally {
      if (isTempFile && sourcePresentationId) { try { Drive.Files.remove(sourcePresentationId); } catch (e) {} }
    }
  }
  
  if (filesToProcess.length > 0) {
    SCRIPT_PROPERTIES.setProperty('filesToProcess', JSON.stringify(filesToProcess));
    SCRIPT_PROPERTIES.setProperty('runCount', (parseInt(properties.runCount, 10) + 1).toString());
    ScriptApp.newTrigger('_continueMergeProcess').timeBased().after(60 * 1000).create();
    Logger.log(`PAUSING: Work remains. The next run has been triggered automatically.`);
  } else {
    Logger.log(`SUCCESS: All files have been merged!`);
    if (mergedPresentation.getSlides().length > 1) { mergedPresentation.getSlides()[0].remove(); }
    Logger.log(`Final presentation: ${mergedPresentation.getUrl()}`);
    _clearPropertiesAndTriggers();
  }
}

function _clearPropertiesAndTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) { ScriptApp.deleteTrigger(trigger); }
  SCRIPT_PROPERTIES.deleteAllProperties();
}

function appendSlideWithRetry(presentation, slide, slideNumber, sourceFileName) {
  const MAX_RETRIES = 3;
  let attempt = 0;
  while (attempt < MAX_RETRIES) {
    try {
      presentation.appendSlide(slide); return;
    } catch (e) {
      attempt++; Utilities.sleep(1000 * attempt);
    }
  }
  Logger.log(`ERROR: Skipped slide #${slideNumber} from "${sourceFileName}" after ${MAX_RETRIES} attempts.`);
}