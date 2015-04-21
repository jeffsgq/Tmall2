<?php
Yii::$enableIncludePath = false;
//发送邮件
Yii::import('application.components.Taobaomail') ;
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
Yii::import('application.components.TaobaoConnector') ;
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' ) ;
include_once (dirname(__FILE__).'/../extensions/PHPExcel/PHPExcel/IOFactory.php');
PHPExcel_CachedObjectStorageFactory::cache_in_memory_serialized;
class CommentSQLCommand extends ConsoleCommand { 
    protected $titleArray = array();
    protected $_evaluateFields = array();
    protected $_tradeField = array();
    protected $_tradeFields = array();
    protected $_tradeFieldsOrders = array();
    protected $PHPWrite = null;
    protected $saveFileName = null;
    protected $fileName = null;
    protected $taobaomail = null;
    public function init(){
        ini_set('memory_limit', '800M');
//        $this->_clearDatabase('comment');//清除数据库数据
        $this->PHPWrite = new PHPExcel();
        $this->fileName = 'trades.xls';
        $this->saveFileName = dirname(__FILE__).'/../../Excel/'.$this->fileName ;
        $this->_evaluateFields = array("tid","oid","nick","result","content");//保证顺序
        $this->_tradeField = array("created", "payment");
        $this->_tradeFields = array("created", "payment","orders");
        $this->_tradeFieldsOrders = array("outer_sku_id","title","oid");
        $this->titleArray = array("tid","oid","nick","result","content","outer_sku_id","title","created","payment");
        fopen($this->saveFileName, "w+");
        //发送邮件
        $this->taobaomail = new Taobaomail();
    }
    public function run($args){
        ob_start();
        $start = date('Y-m-d H:i:s');
       if(count($args)==2){
            $this->_getTraderatesAPIValue($args[0], $args[1]);
        }else if(count($args)==3){
            $this->_getTraderatesAPIValue($args[0], $args[1]);
        }else if(count($args)==0){
            $date=date('Y-m-d');  //当前日期
            $first=1; //$first =1 表示每周星期一为开始日期 0表示每周日为开始日期
            $w=date('w',strtotime($date));  //获取当前周的第几天 周日是 0 周一到周六是 1 - 6
            $now_start=date('Y-m-d',strtotime("$date -".($w ? $w - $first : 6).' days')); //获取本周开始日期，如果$w是0，则表示周日，减去 6 天
            $last_start=date('Y-m-d',strtotime("$now_start - 7 days"));  //上周开始日期
            $last_end=date('Y-m-d',strtotime("$now_start - 1 days"));  //上周结束日期
            $this->_getTraderatesAPIValue($last_start, $last_end);
        }else{
            echo "Please input start date and end date!";
            exit();
        }
        $insertSQL = date('Y-m-d H:i:s');
        $this->_transcation();
        $updateSQL = date('Y-m-d H:i:s');
        $this->_generateExcel();
        $generateExcel = date('Y-m-d H:i:s');
        echo "\nstart time:\t$start\ninsert database:$insertSQL\nupdate database:$updateSQL\ngenerate excel: $generateExcel\n\ttrades.xls\n--------END--------";
        if(!empty($args[2])){
            $this->taobaomail->sendTaobaoMai($this->fileName, $args[2]);
        }else{
            echo "\n**************No email**************";
        }
    }
    public function _generateExcel(){
        $this->_startSaveExcel();
        $count = $this->_getcount();
        $count = $count[0];
        $i = 0;
        $line = 2;
        while($i < $count){
          $j = 100;
          $data =   $this->_getData($i,$j);
          $line = $this->_insertExcel($data, $line);
          $i += $j;
        }
        $this->_endSaveExcel();
    }
    public function _insertExcel($data,$line){
        $currentSheet = $this->PHPWrite->setActiveSheetIndex(0);
        foreach ($data as $_tradeValue){
            $index='A';
            foreach ($this->titleArray as $field){
              $currentSheet->setCellValue(($index++).$line,$_tradeValue[$field]);
            }
            $line++;
        }
        return $line;
    }
    public function _getTraderatesAPIValue($start_date,$end_date){
        $page_no = 0;
        do{
           $_tradesTmall = array();
           $page_no++;
           $_tradeTmall= $this->_connectTmall_Traderates(Yii::app()->params['taobao_api']['accessToken'],$start_date,$end_date,$page_no);
           if(array_key_exists('trade_rates', $_tradeTmall['traderates_get_response'])){
               array_push($_tradesTmall, $_tradeTmall['traderates_get_response']['trade_rates']['trade_rate']);
           }
           $_tradesTmall = $this->_formatFieldsComment($this->_formatArray($_tradesTmall));
           $this->_insertDatabase($_tradesTmall);
           //插入到数据库
        }while($_tradeTmall['traderates_get_response']['has_next']==1);
    }
    public function _getTradeAPIValue($tid) {
        $_itemsTmallAll = array();
        $_itemsTmall = $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'], $tid);
        if (!empty($_itemsTmall)) {
            if (array_key_exists('trade', $_itemsTmall['trade_get_response'])) {
                array_push($_itemsTmallAll, $_itemsTmall['trade_get_response']['trade']);
            } else {
                array_push($_itemsTmallAll, null);
            }
            return $this->_formatFieldsTrade($_itemsTmallAll);
        } else {
            return $_itemsTmall;
        }
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
    public function _formatFieldsComment($_trades){
        $_tradesArray=array();
        foreach ($_trades as $_tradesKey=>$_tradesValue){
            $var_array = array();
            foreach ($this->_evaluateFields as $field) {
                if (array_key_exists($field, $_tradesValue)) {
                     $var_array[$field] = $_tradesValue[$field];
                }
                else {
                     $var_array[$field] = "";
                }
            }
            $var_array['tid'] = number_format($var_array['tid'],0,'','');//获取tid不使用科学技术法
            $var_array['oid'] = number_format($var_array['oid'],0,'','');
            $_tradesArray[] = $var_array;
        }
        return $_tradesArray;
    }
    public function _formatFieldsTrade($_trades){
        $_tradesArray=array();
        foreach ($_trades as $_tradesKey=>$_tradesValue){
            $var_array = array();
            foreach ($this->_tradeFields as $field) {
                if (array_key_exists($field, $_tradesValue)) {
                     $var_array[$field] = $_tradesValue[$field];
                }
                else {
                     $var_array[$field] = "";
                }
            }
            $_tradesArray = $var_array;
        }
        return $_tradesArray;
    }
    
    public function _getTid() {
        $sql_select = "SELECT distinct(tid) FROM comment";
        $connection = Yii::app()->db;
        $command = $connection->createCommand($sql_select);
        $tid = $command->queryAll();
        return $tid;
    }
    public function _transcation(){
        $tids = $this->_formatArray($this->_getTid());
        foreach ($tids as $tidValue){
            $tidTrade = $this->_getTradeAPIValue($tidValue);
            if (!empty($tidTrade)) {
                foreach ($this->_getResultOrders($tidTrade) as $TradeResult){
                    $this->_updateDatabase($TradeResult);
                }
            }     
        }
    }
    public function _getResultOrders($tidTrade){
        $orderValueAll = array();
        $createdPaymentValue = $this->_getResultCreatedPayment($tidTrade);
        $order = $tidTrade['orders'];
        $orders = $this->_formatArray($order);
        foreach ($orders as $orderKey => $orderValue){
            $orderValueTemp = array();
            foreach ($this->_tradeFieldsOrders as $field){
                if (array_key_exists($field, $orderValue)) {
                    $orderValueTemp[$field] = $orderValue[$field];
                } else {
                    $orderValueTemp[$field] = "";
                }
            }
            $orderValueTemp['oid'] = number_format($orderValueTemp['oid'],0,'','');//获取oid不使用科学技术法
            $orderValueAll[] = array_merge($orderValueTemp, $createdPaymentValue);
        }
        return $orderValueAll;
    }
    public function _getResultCreatedPayment($tidTrade){
        $tidTradeResult = array();
        foreach ($this->_tradeField as $field){
            if (array_key_exists($field, $tidTrade)) {
                $tidTradeResult[$field] = $tidTrade[$field];
            } else {
                $tidTradeResult[$field] = "";
            }
        }
        return $tidTradeResult;
    }
    public function _CheckString($str){
            return str_replace("'","''",$str);
        }
    //清除数据库中的数据
    public function _clearDatabase($tblName){
        $sql_clear = "TRUNCATE TABLE ".$tblName;
        $connection= Yii::app()->db;
        $command = $connection->createCommand($sql_clear);
        $command->execute();
    }
    //插入到数据库
    public function _insertDatabase($_trades){
        $connection= Yii::app()->db;//建立数据库连接
        foreach ($_trades as $_tradesKey=>$_tradesValue){
//            print_r($_tradesValue);
            $INSERT = "insert into comment(tid,oid,nick,result,content) values (" . $_tradesValue['tid'] . "," . $_tradesValue['oid'] . ",'" .$this->_CheckString($_tradesValue['nick']). "','" . $_tradesValue['result'] . "','" . $this->_CheckString($_tradesValue['content']) ."')";
            $command2 = $connection->createCommand($INSERT);
            $command2->execute();
        }
    }
    //更新数据
    public function _updateDatabase($TradeResult){
//        print_r($TradeResult);
//        $sql_update = " UPDATE  comment  SET title= '" . $TradeResult['title'] . "' , created = '". $TradeResult['created']. "' , payment = ". $TradeResult['payment'] . "  where oid = ".$TradeResult['oid']."   "; 
        $sql_update = $sql_update = 'UPDATE comment SET title= "' . $this->_CheckString($TradeResult['title']) . '" , created = "'. $TradeResult['created']. '" , payment = "'. $TradeResult['payment'] .'" , outer_sku_id = "'. $TradeResult['outer_sku_id'] . '" where oid = "'.$TradeResult['oid'] . '"'; 
        $connection= Yii::app()->db;//建立数据库连接
        $command2 = $connection->createCommand($sql_update);
        $command2->execute();
    }
     //按每N条取出
    public function _getData($i,$j){
        $sql_get = "SELECT tid,oid,nick,result,outer_sku_id,title,content,created,payment FROM comment LIMIT $i,$j";
        $connection = Yii::app()->db;
        $command4 = $connection->createCommand($sql_get);
        $_Count = $command4->queryAll();
        print_r($_Count);
        return $_Count;
    }
    //获取总行数
     public function _getcount(){
        $sql_count = "SELECT count(id) FROM comment ";
        $connection = Yii::app()->db;
        $command5 = $connection->createCommand($sql_count);
        $_count = $command5->queryAll();
        return $this->_formatArray($_count);
//        return $_count;
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
        $_taobaoConnect=  new TaobaoConnector();
        $_taobaoConnect->__url=Yii::app()->params['taobao_api']['url'] ;
        $_taobaoConnect->__appkey= Yii::app()->params['taobao_api']['appkey'] ;
        $_taobaoConnect->__appsecret= Yii::app()->params['taobao_api']['appsecret'] ;
        $_taobaoConnect->__method= Yii::app()->params['taobao_api']['methods']['evaluate_method'] ;
        $_taobaoConnect->__fields= Yii::app()->params['taobao_api']['fields']['evaluate_field'] ;
        $_items= $_taobaoConnect->connectTaobaoTraderates( $_sessionkey,$start_date,$end_date,$page_no) ;
        if (array_key_exists('error_response',$_items)){
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
            echo "Please input correct date format, like:2015-03-11 2015-04-10 abc@decathlon.com.\nStart date:2015-03-11\nEnd date:2015-04-10\nAnd email address, like:abc@decathlon.com";
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
    private function _connectTmall($_sessionkey, $tid) {
        $_taobaoConnect = new TaobaoConnector();
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
}

