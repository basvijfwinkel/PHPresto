<?php
namespace Presto;

class PHPrestoState
{
  const PRESTO_ERROR   = 'ERROR';
  const PRESTO_SUCCESS = 'SUCCESS';
}

class PHPresto
{
    // variables that can be set by the user
    private $user;       // e.g. 'myuser'
    private $schema;     // e.g. 'default'
    private $catalog;    // e.g 'hive'
    private $userAgent;  // can be empty
    private $url;        // url to connect to e.g http://ec2-123-456-789-012.ap-northeast-1.compute.amazonaws.com:8889/v1/statement
    private $sleepTime;  // number of milisconds to wait before checking the query state. e.g. 500000
    private $maxTries;   // maximum number of retries
    private $currentTry; // current times tried
    private $outputformat; // array or table (first row contains headers)

    // properties that are set by the client
    private $nextUri;          // next query url`
    private $infoUri;          // info query
    private $partialCancelUri; // cancel url
    private $state;            // state of the query
    private $columns;          // columns for the result
    private $headers;          // curl query headers
    private $query;            // presto query
    private $data;             // result data (without column mapping)
    private $error;            // error message

    // Presto states
    const stateUNINITIALIZED = 'UNINITIALIZED';
    const stateRUNNING       = 'RUNNING';
    const stateFINISHED      = 'FINISHED';
    const stateQUEUED        = 'QUEUED';
    const statePLANNING      = 'PLANNING';
    const stateSTARTING      = 'STARTING';
    const stateBLOCKED       = 'BLOCKED';
    const stateFINISHING     = 'FINISHING';
    const stateFAILED        = 'FAILED';

    /**
     * Constructs the presto connection instance
     *
     * @param $config : array with properties to set the state of the client
     *            [ url
     *              catalog
     *              user
    *               schema ]
     */
    public function __construct($config)
    {
       $this->setFromState($config);
    }

    /**
    *   Put all of the object's properties in an ArrayAccess
    *   @return array
    */
    public function serializeConfig()
    {
        $config = [ 'url'              => $this->url,
                    'catalog'          => $this->catalog,
                    'user'             => $this->user,
                    'schema'           => $this->schema,
                    'sleepTime'        => $this->sleepTime,
                    'maxTries'         => $this->maxTries,
                    'currentTry'       => $this->currentTry,
                    'userAgent'        => $this->userAgent,
                    'query'            => $this->query,
                    'nextUri'          => $this->nextUri,
                    'state'            => $this->state,
                    'columns'          => $this->columns,
                    'headers'          => $this->headers,
                    'data'             => $this->data,
                    'error'            => $this->error,
                    'outputformat'     => $this->outputformat,
                    'partialCancelUri' => $this->partialCancelUri];
        return $config;
    }

    /**
    *   get the property value from an array or set a default if the property does not exist
    *   @param $config  : array  : array containing propertyvalues with propertynames as key
    *   @param $name : string : name of the property
    *   @param $default : <any> : any default for when the peroperty is not found in the config ArrayAccess
    *   @return <any> : the property value or the default value
    */
    private function getConfigProperty($config, $name, $default = null)
    {
         return (isset($config[$name])?$config[$name]:$default);
    }

    /**
    *    Set the client state with the config array data
    *    @param $config : array containing the configurable properties
    *
    */
    public function setFromState($config)
    {
       $this->url              = $this->getConfigProperty($config,'url');
       $this->catalog          = $this->getConfigProperty($config,'catalog');
       $this->user             = $this->getConfigProperty($config,'user');
       $this->schema           = $this->getConfigProperty($config,'schema');
       $this->sleepTime        = $this->getConfigProperty($config,'sleepTime',500000);
       $this->maxTries         = $this->getConfigProperty($config,'maxTries',260);
       $this->currentTry       = $this->getConfigProperty($config,'currentTry',0);
       $this->userAgent        = $this->getConfigProperty($config,'userAgent');
       $this->query            = $this->getConfigProperty($config,'query',"");
       $this->nextUri          = $this->getConfigProperty($config,'nextUri',"");
       $this->state            = $this->getConfigProperty($config,'state',self::stateUNINITIALIZED);
       $this->columns          = $this->getConfigProperty($config,'columns');
       $this->headers          = $this->getConfigProperty($config,'headers');
       $this->data             = $this->getConfigProperty($config,'data'  ,[]);
       $this->error            = $this->getConfigProperty($config,'error');
       $this->partialCancelUri = $this->getConfigProperty($config,'partialCancelUri',"");
       $this->outputformat     = $this->getConfigProperty($config,'outputformat',"array");
    }

    /**
     * Return the result data for the query
     * If the 'outputformat' property is set to 'array', an array will be returned.
     * If the 'outputformat' property is set to 'table', an table with a header will be returned
     *
     * @return [PHPrestoState, array]
     */
    public function CombineColumnsAndData()
    {
	      if ($this->state != self::stateFINISHED)
        {
	          return [PHPrestoState::PRESTO_ERROR, "Query has not finished yet. Current state : ".$this->state];
	      }
        if ($this->outputformat == 'array')
        {
            // add the results as an array

            // add colums to result Data
            $result = [];
            foreach($this->data as $entryindex => $entry)
            {
                foreach($entry as $fieldindex => $fielddata)
                {
                    if (isset($this->columns[$fieldindex]->name))
                    {
                        $result[$entryindex][$this->columns[$fieldindex]->name] = $fielddata;
                    }
                    else
                    {
                        return [PHPrestoState::PRESTO_ERROR,'Missing column definition in this->colunms for index :'.$index];
                    }
                }
            }
        }
        else
        {
            // add the results as a table
            $result = [];
            // add headers
            foreach ($this->columns as $column) { $result[0][] = $column->name; }
            // add data
            $result = array_merge($result, $this->data);
        }
        return [PHPrestoState::PRESTO_SUCCESS,$result];
    }

    /**
     * Send the query to the PResto server
     *
     * @param $query
     * @return [PHPrestoState, array]
     *
     */
    public function Query($query)
    {
        $this->query = $query;
        $this->data = [];
        $this->columns = [];

	      //check that no other queries are already running for this object
	      if ($this->state === self::stateRUNNING)
        {
	          return [PHPrestoState::PRESTO_ERROR, "Another query is already running"];
	      }

	      $this->headers = [
	                         "X-Presto-User: ".   $this->user,
	                         "X-Presto-Catalog: ".$this->catalog,
	                         "X-Presto-Schema: ". $this->schema,
	                         //"User-Agent: ".      $this->userAgent
                         ];

        list($postState,$postResult) = $this->postRequest($this->url,$this->headers,$this->query);
        if ($postState == PHPrestoState::PRESTO_ERROR)
        {
            // something went wrong
            return [PHPrestoState::PRESTO_ERROR, $postResult];
        }
        else
        {
            // success posted the query
            $this->GetVarFromResult($postResult);
            return [PHPrestoState::PRESTO_SUCCESS, false];
        }
    }

    /**
    *   Send a post request
    *
    *   @param $url string : url for the post request
    *   @param $headers [array] : post headers
    *   @param $postdata string : post body data
    *   @return [PHPresto state, string] :
    *             in case PHPresto state = PRESTO_ERROR   : [PRESTO_ERROR, error_message ]
    *             in case PHPresto state = PRESTO_SUCCESS : [PRESTO_ERROR, post result data]
    */
    private function postRequest($url, $headers, $postdata)
    {
	      $connect = \curl_init();
	      \curl_setopt($connect,CURLOPT_URL, $url);
	      \curl_setopt($connect,CURLOPT_HTTPHEADER, $headers);
	      \curl_setopt($connect,CURLOPT_RETURNTRANSFER, 1);
	      \curl_setopt($connect,CURLOPT_POST, 1);
	      \curl_setopt($connect,CURLOPT_POSTFIELDS, $postdata);
	      $result = \curl_exec($connect);
	      $httpCode = \curl_getinfo($connect, CURLINFO_HTTP_CODE);
        if($httpCode != "200")
        {
            // some error occured
            return [PHPrestoState::PRESTO_ERROR, 'HTTP ERRROR: '. $httpCode.' Error :'.$result];
        }
        else
        {
            // post request succesfully send
            curl_close($connect);
            return [PHPrestoState::PRESTO_SUCCESS, $result];
        }
    }

    /**
    *   Execute a GET request
    *   @param $url : url for the GET request
    *   @param $headers : headers for the GET request
    *   @return [PHPrestoState, string] : string contains either an error or the result of the GET request
    */
    private function getRequest($url, $headers)
    {
  	      $connect = \curl_init();
  	      \curl_setopt($connect,CURLOPT_URL, $url);
  	      \curl_setopt($connect,CURLOPT_HTTPHEADER, $headers);
  	      \curl_setopt($connect,CURLOPT_RETURNTRANSFER, 1);

  	      $result = \curl_exec($connect);
  	      $httpCode = \curl_getinfo($connect, CURLINFO_HTTP_CODE);
          if(($httpCode != "200") && ($httpCode != "204"))
          {
              return [PHPrestoState::PRESTO_ERROR, 'HTTP ERRROR: '. $httpCode];
          }
          else
          {
              curl_close($connect);
              return [PHPrestoState::PRESTO_SUCCESS, $result];
          }
    }

    /**
     * waits until query was executed
     *
     * @return [bool, array]
     */
    public function WaitQueryExec()
    {
	      while ($this->nextUri)
        {
           list($getState, $result) = $this->getRequest($this->nextUri, $this->headers);
           if ($getState == PHPrestoState::PRESTO_ERROR) { return [PHPrestoState::PRESTO_ERROR, 'Something went wrong while polling for the results :'.var_export($result,1)]; }

           $this->GetVarFromResult($result);
           if(!$this->nextUri)
           {
               $this->currentTry++;
               if ($this->currentTry > $this->maxTries)
               {
                   return [PHPrestoState::PRESTO_ERROR,"Time out waiting for result : waited ".(($this->sleepTime * $this->maxTries)/1000)." seconds."];
               }
    	         usleep($this->sleepTime);
            }
	      }

        // check if the query finished without errors
	      if ($this->state != self::stateFINISHED)
        {
	         return [PHPrestoState::PRESTO_ERROR, $this->error];
        }
        // combine headers and data
        return [PHPrestoState::PRESTO_SUCCESS,$this->CombineColumnsAndData($this->data)];
    }

    /**
    *  Get the state of the current query
    *
    *  @return [PHPrestoState, state]
    */
    public function GetQueryState()
    {
         // query the latest variables
         list($getState,$result) = $this->getRequest($this->nextUri, $this->headers);
         if ($getState == PHPrestoState::PRESTO_ERROR) { return [PHPrestoState::PRESTO_ERROR, $result]; }
         // update the variables
         $this->GetVarFromResult($result);
         return [PHPrestoState::PRESTO_SUCCESS,$this->state];
    }

    /**
    *   Determine if the given state of the query is 'FINISHED'
    *   @param $state string
    *   @return bool : true = FINISHED, false = not finished`
    */
    public function QueryFinished($state)
    {
         return ($state == self::stateFINISHED);
    }

    /**
    *   Determine if the query is still running or getting ready to run
    *   @param $state string
    *   @return bool : true = FINISHED, false = not finished`
    */
    public function QueryRunning($state)
    {
        return (($state == self::stateRUNNING)  ||
                ($state == self::stateQUEUED)   ||
                ($state == self::statePLANNING) ||
                ($state == self::stateSTARTING) ||
                ($state == self::stateFINISHING));
    }

    /**
    *   Get the data from the current query
    *
    *    @return [PHPrestoState, array] : array contains either an error or the resulting dataset
    */
    public function GetData()
    {
        if ($this->state != self::stateFINISHED)
        {
            return [PHPrestoState::PRESTO_ERROR,"Query has not finished yet"];
        }

        return $this->CombineColumnsAndData($this->data);
    }

    /**
     * Provide Information on the query execution
     * The server keeps the information for 15minutes
     * Return the raw JSON message for now
     *
     * @return string
     */
    function GetInfo()
    {
        return $this->getRequest($this->infoUri, $this->headers);
    }

    /**
    *   Extract all relevant data form the json resultset that was returned by the Presto server
    *   @param $result jsonstring : resultset returned by the PResto server
    */
    private function GetVarFromResult($result)
    {
	      /* Retrieve the variables from the JSON answer */
        $decodedJson = json_decode($result);

        if (isset($decodedJson->{'nextUri'}))
        {
            $this->nextUri = $decodedJson->{'nextUri'};} else {$this->nextUri = false;
        }
        if (isset($decodedJson->{'data'}))
        {
           $this->data = array_merge($this->data,$decodedJson->{'data'});
        }
        if (isset($decodedJson->{'infoUri'}))
        {
           $this->infoUri = $decodedJson->{'infoUri'};
        }
        if (isset($decodedJson->{'partialCancelUri'}))
        {
            $this->partialCancelUri = $decodedJson->{'partialCancelUri'};
        }
        if (isset($decodedJson->{'columns'}))
        {
            $this->columns = $decodedJson->{'columns'};
        }
        if (isset($decodedJson->{'stats'}))
        {
    	    $status = $decodedJson->{'stats'};
    	    $this->state = $status->{'state'};
        }
        if (isset($decodedJson->{'error'}))
        {
          $this->error = $decodedJson->{'error'}->{'message'} . ' ' .
                         $decodedJson->{'error'}->{'errorCode'} . ' (' .
                         $decodedJson->{'error'}->{'errorName'} . ' '.
                         $decodedJson->{'error'}->{'errorType'}.')';
        }
	  }

    /**
     * Cancel the current request
     *
     * @return [PHPrestoState, result]
     */
    private function Cancel()
    {
      // check if we have a uri to cancel the current request
	    if (!isset($this->partialCancelUri))
      {
          return [PHPrestoState::PRESTO_ERROR, "No cancel uri provided for this query"];
      }
      // try to cancel the current query
	    return  $this->getRequest($this->partialCancelUri, $this->headers);
    }

    /**
     * Get the error message
     *
     * @return string
     */
     public function getErrorMessage()
    {
         return $this->error;
    }

}
