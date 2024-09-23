## Kørselsgodtgørelse for Skolekørsler - Queue Uploader

This repository retrieves data from an Excel file regarding kørselsgodtgørelser (mileage reimbursement) and uploads it to the **Koerselsgodtgoerelse_egenbefordring** queue using [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

After the data is uploaded, the following robot will handle the queue: [MBU Kørselsgodtgørelse Skolekørsler Queue Handler](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Queue__Handler).

### Process:

1. The robot clears the queue.
2. Deletes old files in the path.
3. Fetches the Excel file with data from SharePoint.
4. Processes the data into a DataFrame.
5. Uploads to the queue.

### Arguments:

- **path**: The file path where the Excel file is saved. This should match the path argument for the 'handler robot'.
- **naeste_agent**: The ID of the 'Næste agent.

### Linear Flow

The linear framework is used when a robot follows a straightforward path from start to finish (A to Z) without fetching jobs from an OpenOrchestrator queue. The flow of the linear framework is illustrated in the following diagram:

![Linear Flow diagram](Robot-Framework.svg)
