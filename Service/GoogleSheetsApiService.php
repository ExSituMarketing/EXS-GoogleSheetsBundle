<?php

namespace EXS\GoogleSheetsBundle\Service;

use Google_Client;
use Google_Service_Sheets;
use Google_Service_Sheets_ValueRange;
use Google_Service_Sheets_BatchUpdateSpreadsheetRequest;
use EXS\GoogleSheetsBundle\Service\GoogleApiClientService;
use EXS\GoogleSheetsBundle\Service\Requests\GoogleSheetsRequests;
use Symfony\Component\Config\Definition\Exception\InvalidConfigurationException;

class GoogleSheetsApiService extends GoogleSheetsRequests
{
    /**
     * Client service
     *
     * @var GoogleApiClientService 
     */
    protected $clientService;
    
    /**
     * Google Service Sheets
     *
     * @var Google_Service_Sheets
     */
    protected $sheetService;
    
    /**
     * Goggle Sreadsheets id
     *
     * @var string
     */
    protected $id;

    /**
     * Initiate the service
     * 
     * @param GoogleApiClientService $clientService
     */
    public function __construct(GoogleApiClientService $clientService)
    {
        $this->clientService = $clientService;
    }

    public function setSheetServices($id)
    {
        $this->id = $id;
        $client = $this->clientService->getClient('offline');   // get api clirent
        $client->setScopes(implode(' ', [Google_Service_Sheets::DRIVE]));   // set permission
        $client = $this->clientService->setClientVerification($client); // set varification          
        $this->sheetService = new Google_Service_Sheets($client);
    }

    public function getGoogleSpreadSheets()
    {
        try {
            $response = $this->sheetService->spreadsheets->get($this->id);
        } catch (\Exception $ex) {
            $response = json_decode($ex->getMessage());
        }
        return $response;
    }

    public function createNewSheet($title = '', $data = [], $header = 0)
    {
        $addNewSheetResponse = $this->addNewSheet($title);
        if ($addNewSheetResponse) {
            $range = $this->getSheetRangeByData($title, $data, $header);
            return $this->InsertSheetData($range, $data);
        }
        false;
    }

    public function addNewSheet($title)
    {
        try {
            $request = $this->getNewSheetRequest($title);
            $requestBody = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
            $requestBody->setRequests($request);            
            $this->sheetService->spreadsheets->batchUpdate($this->id, $requestBody);
            return true;
        } catch (Exception $ex) {
            return false;
        }
    }

    public function getSheetRangeByData($title = '', $data = [], $header = 0, $startCol = 'A')
    {
        if (!is_array($data)) {
            throw new InvalidConfigurationException("Data must be 2 dimentional array");
        }

        $startRow = $header + 1;
        $rows = array_keys($data);
        $numCols = $this->getNumberOfDataCols($rows, $data);
        $endCol = chr(ord($startCol) + ($numCols - 1));
        $endRow = $startRow + (count($rows) - 1);

        return $title . '!' . $startCol . $startRow . ':' . $endCol . $endRow;
    }

    public function getNumberOfDataCols($rows, $data)
    {
        if (isset($data[$rows[0]]) && is_array($data[$rows[0]])) {
            $cols = array_keys($data[$rows[0]]);
            return count($cols);
        }
        throw new InvalidConfigurationException("Data must be 2 dimentional array");
    }

    public function InsertSheetData($range, $data)
    {
        $inputOption = ['valueInputOption' => 'RAW'];
        $requestBody = new Google_Service_Sheets_ValueRange();
        $requestBody->setMajorDimension('ROWS');
        $requestBody->setRange($range);
        $requestBody->setValues($data);
        $response = $this->sheetService->spreadsheets_values->update($this->id, $range, $requestBody, $inputOption);
        return $response->getUpdatedRows();
    }

    public function clearSheetById($sheetId)
    {
        try {
            $request = $this->getClearSheetRequest($sheetId);
            $requestBody = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
            $requestBody->setRequests($request);            
            $this->sheetService->spreadsheets->batchUpdate($this->id, $requestBody);
            return true;
        } catch (Exception $ex) {
            return false;
        }
    }
    
    public function clearSheetByTitle($title)
    {
        $sheetId = $this->getSheetIdByTitle($title);
        if ($sheetId) {
            return $this->clearSheetById($sheetId);
        }
        return false;
    }    

    public function deleteSheetById($sheetId)
    {
        try {
           $request = $this->getDeleteSheetRequest($sheetId);
            $requestBody = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
            $requestBody->setRequests($request);            
            $this->sheetService->spreadsheets->batchUpdate($this->id, $requestBody);
            return true;
        } catch (Exception $ex) {
            return false;
        }
    }

    public function deleteSheetByTitle($title)
    {
        $sheetId = $this->getSheetIdByTitle($title);
        if ($sheetId) {
            return $this->deleteSheetById($sheetId);
        }
        return false;
    }

    public function updateSheet($title, $data, $header)
    {
        $range = $this->getSheetRangeByData($title, $data, $header);
        if ($range) {
            return $this->InsertSheetData($range, $data);
        }
        return false;
    }

    public function getSheetIdByTitle($title)
    {
        $sheets = $this->getGoogleSpreadSheets();
        foreach ($sheets as $key => $sheet) {
            if (isset($sheet->properties->title) && $sheet->properties->title == $title) {
                return $sheet->properties->sheetId;
            }
        }
        return false;
    }
}
