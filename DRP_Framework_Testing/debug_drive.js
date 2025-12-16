function debugDriveFolder() {
  const folderId = "1GT-A2Hkg75uXxQF0FYCKROXW8rBw_XjC";
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  let count = 0;
  const names = [];
  
  Logger.log("Checking folder: " + folder.getName());
  
  while (files.hasNext() && count < 20) {
    const file = files.next();
    names.push(file.getName());
    count++;
  }
  
  Logger.log("Found " + count + " files (first 20):");
  Logger.log(names.join("\n"));
  
  // Test the query specifically
  const query = `'${folderId}' in parents and title contains ' - Partner Dashboard' and trashed = false`;
  const search = DriveApp.searchFiles(query);
  let searchCount = 0;
  while (search.hasNext()) {
    searchCount++;
    search.next();
  }
  Logger.log("Query '" + query + "' found: " + searchCount + " files.");
}
