<?php

namespace EXS\GoogleSheetsBundle\Service;

use Google_Client;
use Google_Service_Sheets;
use Symfony\Component\HttpFoundation\File\Exception\FileNotFoundException;

/**
 * GoogleApiClientService Class
 * 
 * @package EXS\GoogleSheetsBundle\Service
 */
class GoogleApiClientService
{
    /**
     * Application name
     *
     * @var string
     */
    protected $applicationName;

    /**
     * Credential location
     *
     * @var string
     */
    protected $credentials;

    /**
     * User secret location
     *
     * @var string
     */
    protected $clientSecret;

    /**
     * Initiate the service
     * 
     * @param string $applicationName
     * @param string $credentials
     * @param string $clientSecret
     */
    public function __construct($applicationName = '', $credentials = '', $clientSecret = '')
    {
        $this->applicationName = $applicationName;
        $this->credentials = $credentials;
        $this->clientSecret = $clientSecret;
    }

    /**
     * Get the new google api client
     * 
     * @param string $type
     * @return Google_Client
     */
    public function getClient($type = 'offline')
    {
        $client = new Google_Client();
        $client->setApplicationName($this->applicationName);
        $client->setAuthConfig(__DIR__ . $this->clientSecret);
        $client->setAccessType($type);
        return $client;
    }

    /**
     * Validate and set access token 
     * 
     * @param Google_Client $client
     * @return Google_Client
     */
    public function setClientVerification(Google_Client $client)
    {
        $credentialsPath = __DIR__ . $this->credentials;
        $accessToken = $this->getAccessToken($credentialsPath);
        $client->setAccessToken($accessToken);
        return $this->ValidateAccessToken($client, $credentialsPath);
    }

    /**
     * Get access token
     * 
     * @param string $credentialsPath
     * @return array
     */
    public function getAccessToken($credentialsPath = '')
    {
        if (!file_exists($credentialsPath)) {
            throw new FileNotFoundException('Access Token does not exists in ' . $credentialsPath);
        }
        return json_decode(file_get_contents($credentialsPath), true);
    }

    /**
     * Validate access token
     * 
     * @param Google_Client $client
     * @param string $credentialsPath
     * @return Google_Client
     */
    public function ValidateAccessToken(Google_Client $client, $credentialsPath = '')
    {
        if ($client->isAccessTokenExpired() && !empty($credentialsPath)) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            file_put_contents($credentialsPath, json_encode($client->getAccessToken()));
        }
        return $client;
    }

    /**
     * Create the new access token
     * Need to be run on command line manually.
     * 
     * @param Google_Client $client
     * @return string
     */
    public function createNewAccessToken(Google_Client $client)
    {
        $credentialsPath = __DIR__ . $this->credentials;
        $authCode = $this->getVerificationCode($client);
        $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
        try {
            return $this->saveAccessToken($credentialsPath, $accessToken);
        } catch (\Exception $e) {
            return false;
        }
    }

    /**
     * Save the new access token
     * 
     * @param string $credentialsPath
     * @param string $accessToken
     * @return boolean
     */
    public function saveAccessToken($credentialsPath = '', $accessToken = '')
    {
        if (!empty($credentialsPath)) {
            if (!file_exists(dirname($credentialsPath))) {
                mkdir(dirname($credentialsPath), 0700, true);
            }
            file_put_contents($credentialsPath, json_encode($accessToken));
            return true;
        }
        return false;
    }

    /**
     * Get the verification code from command line
     * 
     * @param Google_Client $client
     * @return string
     */
    public function getVerificationCode(Google_Client $client)
    {
        $authUrl = $client->createAuthUrl();
        printf("Open the following link in your browser:\n%s\n", $authUrl);
        print 'Enter verification code: ';
        return trim(fgets(STDIN));
    }

    /**
     * Create the new google sheet api access token
     * 
     * @return boolean
     */
    public function createNewSheetApiAccessToken()
    {
        $client = $this->getClient('offline');
        $client->setScopes(implode(' ', [Google_Service_Sheets::DRIVE]));
        return $this->createNewAccessToken($client);
    }
}
