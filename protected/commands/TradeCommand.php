<?php
Yii::$enableIncludePath = false;
Yii::import('application.components.TaobaoConnectorTraderates') ;
Yii::import('application.components.connectTaobaoTrade') ;
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' ) ;
include_once (dirname(__FILE__).'/../extensions/PHPExcel/PHPExcel/IOFactory.php');
PHPExcel_CachedObjectStorageFactory::cache_in_memory_serialized;
class TradeCommand extends ConsoleCommand { 
    protected $skuArray = null;
    protected $itemArray = null;
    protected $titleArray = null;
    protected $_tradeFields= array();
    protected $_tradeFields2= array();
    protected $PHPWrite = null;
    protected $saveFileName = null;
    public function init(){
        ini_set('memory_limit', '800M');
        $this->PHPWrite = new PHPExcel();
        $this->saveFileName = dirname(__FILE__).'/../../Excel/trades.xls';
        $this->_tradeFields = array("num_iid","tid","nick","result","item_title","content");
        $this->_tradeFields2 = array("created", "payment");
        $this->titleArray = array("num_iid","tid","nick","result","item_title","content","created","payment");
        fopen($this->saveFileName, "w+");
    }
    public function run($args){
        if(count($args)==2){
            $this->_generateExcel($args[0], $args[1]);
        }else if(count($args)==0){
            $date=date('Y-m-d');  //当前日期
            $first=1; //$first =1 表示每周星期一为开始日期 0表示每周日为开始日期
            $w=date('w',strtotime($date));  //获取当前周的第几天 周日是 0 周一到周六是 1 - 6
            $now_start=date('Y-m-d',strtotime("$date -".($w ? $w - $first : 6).' days')); //获取本周开始日期，如果$w是0，则表示周日，减去 6 天
            $last_start=date('Y-m-d',strtotime("$now_start - 7 days"));  //上周开始日期
            $last_end=date('Y-m-d',strtotime("$now_start - 1 days"));  //上周结束日期
            $this->_generateExcel($last_start, $last_end);
        }else{
            echo "Please input start date and end date!";
            exit();
        }
    }
    public function _getTraderatesAPIValue($start_date,$end_date){
        $_tradesTmall= $this->_connectTmall_Traderates(Yii::app()->params['taobao_api']['accessToken'],$start_date,$end_date);
        return $this->_formatArray($_tradesTmall['traderates_get_response']['trade_rates']);
    }
    //获取API属性
    public function _getTradeAPIValue($tid) {
        $_itemsTmallAll = array();
        $_itemsTmall = $this->_connectTmall_Trade(Yii::app()->params['taobao_api']['accessToken'], $tid);
        if (!empty($_itemsTmall)) {
            if (array_key_exists('trade', $_itemsTmall['trade_get_response'])) {
                array_push($_itemsTmallAll, $_itemsTmall['trade_get_response']['trade']);
            } else {
                array_push($_itemsTmallAll, null);
            }
            return $_itemsTmallAll;
        } else {
            return NULL;
        }
    }
    public function _getTrade($tid) {
        $TradeResult = array();
        $_itemTmallAll = $this->_getTradeAPIValue($tid);
        if (!empty($_itemTmallAll)) {
            foreach ($_itemTmallAll as $_firstkey => $_firstvalue) {
                $var_array = array();
                foreach ($this->_tradeFields2 as $field) {
                    if (array_key_exists($field, $_firstvalue)) {
                        $var_array[$field] = $_firstvalue[$field];
                    } else {
                        $var_array[$field] = "";
                    }
                }
                $TradeResult = $var_array;
                unset($var_array);
            }
            unset($_itemTmallAll);
            return $TradeResult;
            
        } else {
            return NULL;
        }
    }
    public function _filterApiValue($start_date,$end_date){
        $_filterResultAll= array();
        $_tradesTmallAll = $this->_getTraderatesAPIValue($start_date,$end_date);
            foreach ($_tradesTmallAll as $_tradesKey=>$_tradesValue){
                $_filterResult1= array();
                foreach ($this->_tradeFields as $field){
                    if(array_key_exists($field,$_tradesValue)){
                        $_filterResult1[$field]= $_tradesValue[$field];
                    }else{
                        $_filterResult1[$field]= "";
                    }
                }
                $tid = number_format($_filterResult1['tid'],0,'','');//获取tid
                $_filterResultAll[] = array_merge($_filterResult1,$this->_getTrade($tid));
                unset($_filterResult1);
            }
            return $_filterResultAll;
    }
    public function _insertExc($start_date,$end_date,$i){
        $currentSheet = $this->PHPWrite->setActiveSheetIndex(0);
        $_filterResult= $this->_filterApiValue($start_date,$end_date);
        foreach ($_filterResult as $_tradeKey=>$_tradeValue){
            $index='A';
            foreach ($this->titleArray as $field){
              $currentSheet->setCellValue(($index++).$i,$_tradeValue[$field]);
              unset($_tradeValue[$field]);
            }
            $i++;
        }
    }
    public function _generateExcel($start_date,$end_date){
        ob_start();
        $this->_startSaveExcel();
        $i = 2;
        $this->_insertExc($start_date,$end_date,$i);
        $this->_endSaveExcel();
        echo "\t".'trades.xls'."\n".'-----------END-----------';
    }
    public function _formatArray($_trades){
        $_tradesArray=array();
        foreach ($_trades as $_firstKey=>$_firstValue){
            foreach ($_firstValue as $_secnodKey=> $_secondValue){
                $_tradesArray[] = $_secondValue;
            }
        }
        return $_tradesArray;
    }
    public function _startSaveExcel(){
        $currentSheet = $this->PHPWrite->setactivesheetindex(0);
        $index = 'A';
        for($i=0;$i<count($this->titleArray);$i++){
            $currentSheet->setCellValue(($index++)."1", $this->titleArray[$i]);    
        }
    }
    public function _endSaveExcel(){ 
        if(!is_writable($this->saveFileName)){
            echo 'Can not Write';
            exit();
        }
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename='.$this->saveFileName);
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($this->PHPWrite,'Excel5');  
        $objWriter->save($this->saveFileName);
    }
    private function _connectTmall_Traderates($_sessionkey,$start_date,$end_date){
        $_taobaoConnect=  new TaobaoConnectorTraderates();
        $_taobaoConnect->__url=Yii::app()->params['taobao_api']['url'] ;
        $_taobaoConnect->__appkey= Yii::app()->params['taobao_api']['appkey'] ;
        $_taobaoConnect->__appsecret= Yii::app()->params['taobao_api']['appsecret'] ;
        $_taobaoConnect->__method= Yii::app()->params['taobao_api']['method6'] ;
        $_taobaoConnect->__fields= Yii::app()->params['taobao_api']['fields6'] ;
        $_items= $_taobaoConnect->connectTaobaoTraderates( $_sessionkey,$start_date,$end_date) ;
        if (array_key_exists('error_response',$_items)){
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
            echo "Please input correct date format, like:2015-03-11 2015-04-10.\nstart date:2015-03-11\nend date:2015-04-10";
            exit();
        }
        if (array_key_exists('traderates_get_response',$_items)){
                if (array_key_exists('trade_rates',$_items['traderates_get_response'])){
                    return $_items;
                }else{
                    Yii::log('traderates_get_response not exists data from ' .$start_date. ' to '. $end_date, 'error', 'system.fail');
                    echo 'From '.$start_date.' to '.$end_date.' not exists data!';
                    exit();
                }   
        }else{
            Yii::log('traderates_get_response not exists data from ' .$start_date. ' to '. $end_date, 'error', 'system.fail');
            exit();
        }
    }
    private function _connectTmall_Trade($_sessionkey, $tid) {
       $_taobaoConnect = new TaobaoConnectorTrade();
       $_taobaoConnect->__url = Yii::app()->params['taobao_api']['url'];
       $_taobaoConnect->__appkey = Yii::app()->params['taobao_api']['appkey'];
       $_taobaoConnect->__appsecret = Yii::app()->params['taobao_api']['appsecret'];
       $_taobaoConnect->__method = Yii::app()->params['taobao_api']['method7'];
       $_taobaoConnect->__fields = Yii::app()->params['taobao_api']['fields7'];
       $_items = $_taobaoConnect->connectTaobaoTrade($_sessionkey, $tid);
       if (array_key_exists('error_response', $_items)) {
           Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
           return NULL;
       }
       if (array_key_exists('trade_get_response', $_items)) {
           if (!empty($_items)) {
               return $_items;
           } else {
               Yii::log('No data tid', 'error', 'system.fail');
               return NULL;
           }
       } else {
           return NULL;
       }
    }
}

