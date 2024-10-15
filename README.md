## Kørselsgodtgørelse for Skolekørsler - Queue Uploader

This robot is part of the 'MBU Koerselsgodtgoerelse Skolekoersler' process.

#### Process Overview

The process consists of four main robots working in sequence:

1. **Create Excel and Upload to SharePoint**:  
   The first robot retrieves and exports weekly 'Egenbefordring' data from a database to an Excel file, which is then uploaded to SharePoint at the following location: `MBU - RPA - Egenbefordring/Dokumenter/General`. Once the file is processed, personnel will move it to `MBU - RPA - Egenbefordring/Dokumenter/General/Til udbetaling`. Run it with the 'Single Trigger' or with the Scheduled Trigger'.

2. **Queue Uploader**:  
   The second robot retrieves data from the Excel file and uploads it to the **Koerselsgodtgoerelse_egenbefordring** queue using [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator). Run it with the 'Single Trigger'.

3. **Queue Handler**:  
   The third robot, triggered by the Queue Trigger in OpenOrchestrator, processes the queue elements by creating tickets in OPUS.

4. **Update SharePoint**:  
   The fourth robot cleans and updates the files in SharePoint by uploading the updated Excel file and attachments of any failed elements. Run it with the 'Single Trigger'.

### The Queue Uploader Process

This robot retrieves data from an Excel file regarding kørselsgodtgørelser and uploads it to the **Koerselsgodtgoerelse_egenbefordring** queue.

Steps:
1. Clears the queue.
2. Deletes old files in the specified path.
3. Fetches the Excel file with data from SharePoint.
4. Processes the data into a DataFrame.
5. Uploads the data to the queue.

### Process and Related Robots

1. **Create Excel & Upload to SharePoint**: [Create Excel & Upload To SharePoint](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Dan_Excel_Upload_Til_SharePoint)
2. **Queue Uploader** (This Robot)
3. **Queue Handler**: [Queue Handler](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Queue__Handler)
4. **Update SharePoint**: [Update Sharepoint](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Update_Sharepoint)

### Arguments

- **path**: The file path where the Excel file is saved. This should match the path argument used in the 'Queue Handler' robot.
- **naeste_agent**: The ID of the 'Næste agent'.
