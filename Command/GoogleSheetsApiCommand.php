<?php

namespace EXS\GoogleSheetsBundle\Command;

use Symfony\Bundle\FrameworkBundle\Command\ContainerAwareCommand;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;

class GoogleSheetsApiCommand extends ContainerAwareCommand
{
    private $command = 'googlesheets:execute';

    protected function configure()
    {
        $this->setName($this->command)
        ->addOption('function', null, InputOption::VALUE_OPTIONAL,'sheets api function to be executed')
        ->addOption('title', null, InputOption::VALUE_OPTIONAL,'sheet title in string')
        ->addOption('id', null, InputOption::VALUE_OPTIONAL,'spreadsheets id in integer', 0)
        ->addOption('header', null, InputOption::VALUE_OPTIONAL,'number of rows for the header', 0)
        ->addOption('data', null, InputOption::VALUE_OPTIONAL,'grid data in 2 dimesional array');
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $function = $input->getOption('function');
        $id = $input->getOption('id');
        $sheetTitle = $input->getOption('title');
        $data = json_decode($input->getOption('data'));
        $header = $input->getOption('header');
    
        $clientService = $this->getContainer()->get('exs_googles_heets.api_client_service');
        $service = $this->getContainer()->get('exs_google_sheets.sheets_service');

        $response = 'no action has been made';
        if($function == 'token') {
            $response = $clientService->createNewSheetApiAccessToken();
        } else {
            $service->setSheetServices($id);
        }        
        
        if($function == 'get') {
            $response = $service->getGoogleSpreadSheets();         
        } elseif($function == 'create') {
            $response = $service->createNewSheet($sheetTitle, $data, $header);           
        } elseif($function == 'update') {
            $response = $service->updateSheet($sheetTitle, $data, $header);
        } elseif($function == 'clear') {
            $response = $service->clearSheetByTitle($sheetTitle);     
        } elseif($function == 'delete') {
            $response = $service->deleteSheetByTitle($sheetTitle);        
        }
        
        $output->writeln($response);
    }    
}
