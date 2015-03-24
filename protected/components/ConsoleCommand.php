<?php
/**
 * Extended yii console command
 * @author: radzserg
 * @date: 11.04.11
 */
require_once(dirname(__FILE__) . '/../../../framework/console/CConsoleCommand.php') ;
class ConsoleCommand extends CConsoleCommand {
 
    const VERBOSE_ERROR = 'error';
    const VERBOSE_INFO = 'info';
    const VERBOSE_SYSTEM = 'system';
 
    public $verbose;
    public $_last_id= null ;
    public $_action= null ;
    private $_lockFile ;
     
    // if false this means that multiple scripts can work simultaneously
 
    protected $_isSingletonScript = false;
 
    // calculate time execution time
 
    protected $_timeStart;
    
    protected $_function;
    
    protected function _verbose($message, $level=null, $type=null, $request=null) {
        $this->_callSql( $request ) ;
        if (!$this->verbose) { return ; }
 
        $level = (int)$level ;
        $indent= str_repeat("\t", $level);
        if( $type == self::VERBOSE_ERROR) {
            // message in red
            $message = "\033[31;1m" . $message . "\033[0m\n";
        } elseif( $type== self::VERBOSE_INFO ) {
            // message in green
            $message = "\033[32;1m" . $message . "\033[0m\n";
        } elseif( $type== self::VERBOSE_SYSTEM) {
            $message = "\033[33;1m" . $message . "\033[0m\n";
        }
 
        echo $indent . date('H:i:s ') . $message . "\n";
    }
 
    protected function beforeAction($action, $params) {
        $this->_action= str_replace( 'Command', '', $action ) ;
        $this->_verbose("Start execution of " . get_class($this), null, self::VERBOSE_SYSTEM ) ;
        $this->_timeStart= $this->_microtimeFloat();
        if( $this->_isSingletonScript ) {
            $lockDir= Yii::getPathOfAlias('application.commands.lock');
            if( !is_dir( $lockDir ) ) {
                mkdir( $lockDir ) ;
            }
            $filePath= $lockDir . '/' . get_class($this) . '.lock';
            $this->_lockFile= fopen($filePath, "w");
            if( !flock($this->_lockFile, LOCK_EX | LOCK_NB ) ) {
                $this->_verbose("Another instance of this script is running") ;
                return false ;
            }
        }
        return true ;
    }
 
    protected function afterAction($action, $params, $exitCode=0 ) {
        $this->_action= str_replace( 'Command', '', $action ) ;
        if ($this->_lockFile) { flock($this->_lockFile, LOCK_UN ) ; }
        $time = round($this->_microtimeFloat() - $this->_timeStart, 2);
        $this->_verbose("End (time: {$time} seconds)", null, self::VERBOSE_SYSTEM, array('status'=> 'ending' ) ) ;
    }
    /**
     * get help with global options
     * @see CConsoleCommand::getHelp()
     * @return string
     */
    public function getHelp() {
        $help = parent::getHelp();
        $global_options = $this->getGlobalOptions();
        if (!empty($global_options)) {
            $help .= PHP_EOL . 'Global options:';
            foreach ($global_options as $name => $value) {
                $help .= PHP_EOL . '    [' . $name . '=' . $value . ']';
            }
        }
        return $help;
    }

    /**
     * collect global options
     * @return array
     */
    protected function getGlobalOptions() {
        $options = array();
        $refl = new ReflectionClass($this);
        $properties = $refl->getProperties(ReflectionProperty::IS_PUBLIC);
        foreach ($properties as $property) {
            if ($property->getName() != 'defaultAction') {
                $options[$property->getName()] = $property->getValue($this);
            }
        }
        return $options;
    }

    /**
     * show this information
     *
     * @return void
     */
    public function actionHelp() {
        $this->printf("Info: " . $this->getHelp());
    }

    /**
     * printf with line break
     *
     * @see printf()
     * @return void
     */
    protected function printf() {
        $args= func_get_args(); // PHP 5.2 workaround
        call_user_func_array('printf', $args ) ;
        printf(PHP_EOL);
    }
    
 
    private function _microtimeFloat() {
        list($usec, $sec) = explode(" ", microtime());
        return ((float)$usec + (float)$sec);
    }
    
    private function _callSql( $request ) {
        $command= Yii::app()->db->createCommand() ;
        if( is_null( $request ) ) {
            $_params= '' ;
            if(!empty($_SERVER['argv'])){
                foreach($_SERVER['argv'] as $keyParam=> $valueParam ) {
                    if( $keyParam> 0 ) { $_params.= ' ' . $valueParam ; }
                }
            }else{
                echo $this->_function;
            }
            $command->insert(
                            'commands', 
                            array( 
                                'cmd_execute'=> $_params,
                                'cmd_name'=> $this->_action, 
                                'cmd_starting'=> date('Y-m-d H:i:s') 
                            ) 
            ) ;
            $this->_last_id= Yii::app()->db->getLastInsertID() ;
        }
        else {
            if( (int)$this->_last_id!= 0 ) {
                $command->update(
                            'commands', 
                            array( 
                                'cmd_ending'=> date('Y-m-d H:i:s') 
                            ), 
                            'command_id=:id', 
                            array(':id'=> $this->_last_id ) 
                ) ;
            }
        }
    }
//echo '---' ; var_dump( $action, $params ) ; exit() ;
//    Yii::import('application.commands.OrdersMissingCommand') ;
//    $method= new ReflectionMethod('OrdersMissingCommand', 'actionOrdersMissing') ;
//    echo '<pre>';
//    print_r( $method->getParameters() ); // prints 'A' 
//exit();        
//echo '---' ; var_dump( $action, $params ) ; exit() ;         
        
    
    
    
}