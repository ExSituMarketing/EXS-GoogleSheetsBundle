<?php

namespace EXS\GoogleSheetsBundle\Service\Requests;

class GoogleSheetsRequests
{
    /**
     * Get add new sheet request 
     * 
     * @param string $title
     * @param string $type
     * @param array $gidProperties
     * @return array
     */
    public function getNewSheetRequest($title = '', $type = 'GRID', $gidProperties = null)
    {
        return [
            'addSheet' => [
                'properties' => [
                    "title" => $title,
                    "sheetType" => $type,
                    "gridProperties" => $gidProperties
                ]
            ]
        ];
    }

    /**
     * Get clear sheet request
     * 
     * @param int $sheetId
     * @param string $fieldOption
     * @return array
     */
    public function getClearSheetRequest($sheetId, $fieldOption = 'userEnteredValue')
    {
        return [
            'updateCells' => [
                'range' => [
                    "sheetId" => $sheetId
                ],
                "fields" => $fieldOption
            ]
        ];
    }

    /**
     * Get delete sheet request
     * 
     * @param int $sheetId
     * @return array
     */
    public function getDeleteSheetRequest($sheetId)
    {
        return [
            'deleteSheet' => [
                "sheetId" => $sheetId
            ]
        ];
    }
}
