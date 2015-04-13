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
    protected $_tradeFields2_order= array();
    protected $PHPWrite = null;
    protected $saveFileName = null;
    public function init(){
        ini_set('memory_limit', '800M');
        $this->PHPWrite = new PHPExcel();
        $this->saveFileName = dirname(__FILE__).'/../../Excel/trades.xls';
        $this->_tradeFields = array("tid","oid","nick","result","content");
        $this->_tradeFields2 = array("created", "payment");
        $this->_tradeFields2_order = array("title", "outer_sku_id", "oid");
        $this->titleArray = array("tid","oid","nick","result","outer_sku_id","title","content","created","payment");
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
        $page_no = 0;
        $_tradesTmall = array();
        do{
           $page_no++;
           $_tradeTmall= $this->_connectTmall_Traderates(Yii::app()->params['taobao_api']['accessToken'],$start_date,$end_date,$page_no);
           if(array_key_exists('trade_rates', $_tradeTmall['traderates_get_response'])){
               array_push($_tradesTmall, $_tradeTmall['traderates_get_response']['trade_rates']['trade_rate']);
           }
        }while($_tradeTmall['traderates_get_response']['has_next']==1);
        return $this->_formatArray($_tradesTmall);
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
     public function _getTradeAPIValue_orders($tid) {
        $_itemsTmallAll_orders = array();
        $_itemsTmall = $this->_connectTmall_orders(Yii::app()->params['taobao_api']['accessToken'], $tid);
        if (!empty($_itemsTmall)) {
            if (array_key_exists('trade', $_itemsTmall['trade_get_response'])) {
                array_push($_itemsTmallAll_orders, $_itemsTmall['trade_get_response']['trade']['orders']['order']);
            } else {
                array_push($_itemsTmallAll_orders, null);
            }
            return $_itemsTmallAll_orders;
        } else {

            return $_itemsTmall;
        }
    }
    public function _getTrade($tid,$oid) {
        $TradeResult = array();
        $_itemTmallAll = $this->_getTradeAPIValue($tid);
        $_itemTmallAll_orders = $this->_getTradeAPIValue_orders($tid);
        if (!empty($_itemTmallAll)) {
           foreach ($_itemTmallAll as $_firstkey => $_firstvalue) {
               foreach ($this->_tradeFields2 as $field) {
                   if (array_key_exists($field, $_firstvalue)) {
                       $var_array[$field] = $_firstvalue[$field];
                   }
                   else {
                    $var_array[$field] = "";
                   }
               }
           } 
       }
       else{
           return NULL;
       }
       if (!empty($_itemTmallAll_orders)) {
            for ($i = 1; $i <= count($_itemTmallAll_orders); $i++) {
                foreach ($_itemTmallAll_orders[$i-1] as $_firstkey => $_firstvalue) {
                    foreach ($this->_tradeFields2_order as $field) {
                        if (array_key_exists($field, $_firstvalue)) {
                            $var_array_orders[$field] = $_firstvalue[$field];  
                        }
                        else
                        {
                         $var_array_orders[$field] = "";
                        }
                    }
                }
                if($var_array_orders['oid']==$oid){
                        $TradeResult = array_merge($var_array,$var_array_orders);
                }else{
                    $var_array_orders = "";
                    continue;
                }
            }
            return  $TradeResult;       
        }else{
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
                $tid = number_format($_filterResult1['tid'],0,'','');//获取tid不使用科学技术法
                $oid = number_format($_filterResult1['oid'],0,'','');//获取oid不使用科学技术法
                $_filterResultAll[] = array_merge($_filterResult1,$this->_getTrade($tid,$oid));
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
    private function _connectTmall_Traderates($_sessionkey,$start_date,$end_date,$page_no){
        $_taobaoConnect=  new TaobaoConnectorTraderates();
        $_taobaoConnect->__url=Yii::app()->params['taobao_api']['url'] ;
        $_taobaoConnect->__appkey= Yii::app()->params['taobao_api']['appkey'] ;
        $_taobaoConnect->__appsecret= Yii::app()->params['taobao_api']['appsecret'] ;
        $_taobaoConnect->__method= Yii::app()->params['taobao_api']['methods']['evaluate_method'] ;
        $_taobaoConnect->__fields= Yii::app()->params['taobao_api']['fields']['evaluate_field'] ;
        $_items= $_taobaoConnect->connectTaobaoTraderates( $_sessionkey,$start_date,$end_date,$page_no) ;
        if (array_key_exists('error_response',$_items)){
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
            echo "Please input correct date format, like:2015-03-11 2015-04-10.\nStart date:2015-03-11\nEnd date:2015-04-10";
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
    //获取created,payment
    private function _connectTmall_Trade($_sessionkey, $tid) {
        $_taobaoConnect = new TaobaoConnectorTrade();
        $_taobaoConnect->__url = Yii::app()->params['taobao_api']['url'];
        $_taobaoConnect->__appkey = Yii::app()->params['taobao_api']['appkey'];
        $_taobaoConnect->__appsecret = Yii::app()->params['taobao_api']['appsecret'];
        $_taobaoConnect->__method = Yii::app()->params['taobao_api']['methods']['trade_method'];
        $_taobaoConnect->__fields = Yii::app()->params['taobao_api']['fields']['trade_field'];
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
    //获取orders.title,orders.outer_sku_id,orders.oid
    private function _connectTmall_orders($_sessionkey, $tid) {
        $_taobaoConnect = new TaobaoConnectorTrade();
        $_taobaoConnect->__url = Yii::app()->params['taobao_api']['url'];
        $_taobaoConnect->__appkey = Yii::app()->params['taobao_api']['appkey'];
        $_taobaoConnect->__appsecret = Yii::app()->params['taobao_api']['appsecret'];
        $_taobaoConnect->__method = Yii::app()->params['taobao_api']['methods']['trade_method'];
        $_taobaoConnect->__fields = Yii::app()->params['taobao_api']['fields']['trade_field_orders'];
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

