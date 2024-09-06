## Kørselsgodtgørelse for skolekørsler - Queue Uploader

This repository retrieves data from an Excel file regarding kørselsgodtgørelser and uploads it to the "Koerselsgodtgoerelse_egenbefordring" queue using [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

### Linear Flow

The linear framework is used when a robot is just going from A to Z without fetching jobs from an
OpenOrchestrator queue.
The flow of the linear framework is sketched up in the following illustration:

![Linear Flow diagram](Robot-Framework.svg)
