<?php
/*
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at:
http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.*/

use \Presto\PHPrestoState;
use \Presto\PHPresto;

require_once(__DIR__ . '/../src/PHPresto.php');

// try to load the connection settings from the presto.ini file
$config = false;
if (file_exists('presto.ini'))
{
    // read config from ini FAILED
    $config = parse_ini_file("presto.ini");
}
// check if we able to load the config settings form the ini file
if (!$config)
{
    // no presto.ini file or loading the file failed
    $config = [ 'url'     => 'http://my.presto.server.domain.com:8889/v1/statement',
                'catalog' => 'hive',
                'user'    => 'my_user_name',
                'schema'  => 'default',
                //'outputformat' => 'table'
              ];
}

// create a new client object
$presto = new PHPresto($config);

// your sql request
$query = "show schemas";

try
{
    // try to send the query to the Presto server
    list($queryState,$result) = $presto->Query($query);
}
catch (\Exception $e)
{
    var_dump($e);die();
}

// check if the query was succesfull send
if ($queryState != PHPrestoState::PRESTO_ERROR)
{
    // get all required info from the client and destroy the client
    // in order to simulate an asynchronous connection to the Presto server.
    // the stateArray can be serialized to a string and stored somewhere
    $stateArray = $presto->serializeConfig();
    $presto = null;

    // start polling the server
    $times = 100;
    while($times >= 0)
    {
        // create a new client with the stateArray properties.
        // these properties will be updated when polling
        $presto = new PHPresto($stateArray);
        // get the current query state
        list($prestoState, $querystate) = $presto->GetQueryState();
        if ($prestoState == PHPrestoState::PRESTO_ERROR)
        {
              // some error occured while retrieving the state
              $prestoState = PHPRestoState::PRESTO_ERROR;
              $result = 'Error while retrieving query state:'.var_export($querystate,1);
              break;
        }
        // check if the query finished
        if ($presto->QueryFinished($querystate))
        {
           // get the data because the query is finished
           list($prestoState, $result) = $presto->GetData();
           break;
        }
        else if (!$presto->QueryRunning($querystate))
        {
           // the query if not finished and not running : unexpected state
           $prestoState = PHPRestoState::PRESTO_ERROR;
           $result = "The query did not finish as expected : ".$presto->getErrorMessage();
           break;
        }

        // if we end up here, the query is still running
        // destroy the client to simulate an asynchronous connection
        $stateArray = $presto->serializeConfig();
        $presto = null;

        // wait for a second before polling again
        sleep(1);
        $times--;
        if ($times == 0) { $prestoState = PHPRestoState::PRESTO_ERROR; $result = "Timeout in async version, script not finished yet"; }
    }

    // show the result if we didn't end up with an error
    if ($prestoState != PHPRestoState::PRESTO_ERROR)
    {
       // display the result
       echo("The result for the query : ");
       echo(var_export($result,1));
    }
    else
    {
        // something went wrong while waiting for the query
        echo("Something went wrong while waiting for the query : ");
        var_export($result);
    }
}
else
{
    // something went wrong while sending the query
    echo("Something went wrong sending the query.");
    var_export($result);
}
