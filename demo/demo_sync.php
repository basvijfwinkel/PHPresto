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
limitations under the License.
*/

use \PHPresto\PHPrestoState;
use \PHPresto\PHPresto;

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

// create an instance of the client
$presto = new PHPresto($config);

// your sql request
$query = "show schemas";

try
{
    // try to send the query to PResto
    list($queryState,$result) = $presto->Query($query);
}
catch (\Exception $e)
{
    var_dump($e);die();
}

// check if Presto returned an error
if ($queryState != PHPrestoState::PRESTO_ERROR)
{
    // poll the Presto server to see if it finished executing the query
    // and retrieve the result
    list($resultState, $result) = $presto->WaitQueryExec();
    if ($resultState != PHPrestoState::PRESTO_ERROR)
    {
       // display the result;
       echo("The result for the query : ");
       echo(var_export($result,1));
    }
    else
    {
        // some error occured while waiting for the result or parsing the data
        echo("Something went wrong while waiting for the query : ");
        var_export($result);
    }
}
else
{
    // something went wrong with sending the query
    echo("Something went wrong sending the query.<br>");
    var_export($result);
}
