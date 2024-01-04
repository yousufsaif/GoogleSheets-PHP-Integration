<?php
/**
 * Project: Google Sheets PHP Integration
 * Author: Yousuf Saif
 * GitHub: https://github.com/yousufsaif
 *
 * This PHP script is designed for integrating with Google Sheets API.
 * It allows for operations such as reading, updating, appending, and clearing data.
 * Created and maintained by Yousuf Saif.
 */

 
declare(strict_types=1);

require __DIR__ . "/vendor/autoload.php";

class GoogleSheetsService
{
    private $service;
    private $spreadsheetId;


    /**
     * Constructs the Google Sheets Service object.
     *
     * Initializes the Google Sheets API service and sets the spreadsheet ID for further operations.
     * 
     * @param Google_Client $client The Google Client object, pre-configured with API credentials and settings.
     * @param string $spreadsheetId The ID of the Google Spreadsheet to be accessed or manipulated.
     */
    public function __construct(Google_Client $client, string $spreadsheetId)
    {
        $this->service = new Google_Service_Sheets($client);
        $this->spreadsheetId = $spreadsheetId;
    }


    /**
     * Retrieves data from a specified range in the Google Spreadsheet.
     *
     * Fetches and returns the values from a specific range within the spreadsheet.
     * Returns an empty array if no data is found.
     * 
     * @param string $range A string representing the range to retrieve in A1 notation.
     * @return array An array of spreadsheet values. Each element is a row, represented as an array.
     */
    public function getData(string $range): array
    {
        $response = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
        return $response->getValues() ?? [];
    }


    /**
     * Updates data in a specified range of the Google Spreadsheet.
     *
     * Takes a range and an array of new values, then updates the spreadsheet at the specified range.
     * 
     * @param string $range A string representing the range to update in A1 notation.
     * @param array $newValues An array of new values to be updated in the specified range.
     */
    public function updateData(string $range, array $newValues): void
    {
        $values = [$newValues];
        $body = new Google_Service_Sheets_ValueRange(['values' => $values]);
        $params = ['valueInputOption' => 'RAW'];
        $this->service->spreadsheets_values->update($this->spreadsheetId, $range, $body, $params);
    }


    /**
     * Appends data to the end of the specified range in the Google Spreadsheet.
     *
     * Adds new rows of data to the end of a specified range.
     * 
     * @param string $range A string representing the range to append to in A1 notation.
     * @param array $newValues An array of new values to be appended.
     */
    public function appendData(string $range, array $newValues): void
    {
        $values = [$newValues];
        $body = new Google_Service_Sheets_ValueRange(['values' => $values]);
        $params = ['valueInputOption' => 'RAW'];
        $insert = ["insertDataOption" => "INSERT_ROWS"];
        $this->service->spreadsheets_values->append($this->spreadsheetId, $range, $body, $params, $insert);
    }


    /**
     * Clears data from a specified range in the Google Spreadsheet.
     *
     * Removes all the data from the specified range, leaving the cells empty.
     * 
     * @param string $range A string representing the range to be cleared in A1 notation.
     */
    public function clearData(string $range): void
    {
        $this->service->spreadsheets_values->clear($this->spreadsheetId, $range, new Google_Service_Sheets_ClearValuesRequest());
    }
}

$client = new \Google_Client();
$client->setApplicationName('SheetsViaPHP');
$client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
$client->setAccessType('offline');
$client->setAuthConfig(__DIR__ . '/credentials.json');

$spreadsheetId = "your_spreadsheet_id_here"; // Replace with your actual spreadsheet ID
$googleSheetsService = new GoogleSheetsService($client, $spreadsheetId);

// Example usage
$data = $googleSheetsService->getData("Sheet1!A1:C10");
foreach ($data as $row) {
    echo implode("\t", $row) . "<br>";
}

// Example to update data
$googleSheetsService->updateData("Sheet1!A2", ["New Value1", "New Value2"]);

// Example to append data
$googleSheetsService->appendData("Sheet1", ["New Row1", "New Row2"]);

// Example to clear data
$googleSheetsService->clearData("Sheet1!A2:B2");
