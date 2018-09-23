<?php

namespace App;

use GuzzleHttp\Client;

class ServerData {

    private $url;
    private $username;
    private $password;
    private $query;

    public function __construct(string $url, string $username, string $password, array $query)
    {
        $this->url = $url;
        $this->username = $username;
        $this->password = $password;
        $this->query = $query;
    }

    // get data from rest api
    public function get()
    {
        $client = new Client();

        $response = $client->get($this->url, [
            'auth' => [
                $this->username,
                $this->password
            ],
            'query' => $this->query
        ]);

        return json_decode($response->getBody());
    }
}