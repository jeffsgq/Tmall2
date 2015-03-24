<?php
Yii::import('application.components.TaobaoConnector') ;
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' );
//我添加的代码
Yii::$enableIncludePath = false; 
include_once (dirname(__FILE__).'/../extensions/PHPExcel/PHPExcel/IOFactory.php');
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
class TestCommand extends ConsoleCommand {
    
    protected $_className= null ;
    protected $_startDate= null ;
    protected $_endDate= null ;
    protected $_nowTime=null ;
    
    public function init(){
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
        $this->_nowTime= Yii::app()->params['today_time'];
        $this->_startDate= date("Y-m-d H:i:s",strtotime($this->_nowTime . Yii::app()->params['modify_time']['less_3_weeks']));
        $this->_endDate= date("Y-m-d H:i:s",strtotime($this->_nowTime . Yii::app()->params['modify_time']['less_10_mins']));//修改less_days
    }
    
    public function run($args){
        $_page=0;
        $_ordersTmall= array();
        do{
            $_page++;
            $_pageOrdersTmall= $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'],'true',$_page);
            if (array_key_exists('trades',$_pageOrdersTmall['trades_sold_get_response'])){
                array_push($_ordersTmall, $_pageOrdersTmall['trades_sold_get_response']['trades']['trade']);
            }
        }while($_pageOrdersTmall['trades_sold_get_response']['has_next']==1);
        //调用输出方法
        $this->_Print($_ordersTmall);
        print_r($this->_jsonNbFormat($this->_formatArray($_ordersTmall)));
    }
    
    //循环输出并保存在Excel中
    public function _Print($_ordersTmall){
        ob_start();
        //调用PHPExcel
        $PHPExcel = new PHPExcel();
        $PHPExcel->setActiveSheetIndex(0)
            //向Excel中添加数据
            ->setCellValue('A1', 'created')
            ->setCellValue('B1', 'pay_time')
            ->setCellValue('C1', 'status')
            ->setCellValue('D1', 'tid');
        //设置单元格大小
        $PHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(23);
        $PHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(23);
        $PHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(23);
        $PHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(23);
        //循环
        $_orders = $this->_formatArray($_ordersTmall);
        $_orders2 = $this->_jsonNbFormat($_orders);
        $i=2;
        //循环输出
        foreach ($_orders2 as $_firstKey=>$_firstValue){
            $j = 65;
            foreach ($_firstValue as $_secnodKey=> $_secondValue){
                //写入
                $PHPExcel->setActiveSheetIndex(0)->setCellValue(chr($j).$i, $_secondValue);
                $j = $j+1;
            }
            $i = $i + 1;
        }
        //定义Excel
        $PHPExcel->setActiveSheetIndex(0);
        $name = "D:\php.xls";//不能将文件夹放在不可执行的路径下
        if(!is_writable($name)){
            echo 'Can not Write';
            return ;
        }
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename='.$name);
        header('Cache-Control: max-age=0');
        //创建文件使用Excel2003版本
        $objWriter = PHPExcel_IOFactory::createWriter($PHPExcel,'Excel5');  
        $objWriter->save($name);
        echo 'END';
    }
   
    private function _formatArray($_orders){
        $_ordersArray=array();
        foreach ($_orders as $_firstKey=>$_firstValue){
            foreach ($_firstValue as $_secnodKey=> $_secondValue){
                $_ordersArray[] = $_secondValue;
            }
        }
        return $_ordersArray;
    }
    
    private function _jsonNbFormat($_orders){
        foreach ($_orders as $key=>$val){
            $_orders[$key]['tid']= number_format($val['tid'],0,'','');
        }
        return $_orders;
    }
    
    private function _connectTmall($_sessionkey,$hasNext,$pageNo=""){
        ob_start();
        //此处Excel写入
        $fileName = 'D:\php2.xls';
        if(!is_readable($fileName)){//判断文件是否可读
            throw new Exception("No such file or directory"); 
        }
        $PHPReader = new PHPExcel_Reader_Excel2007();
        if(!$PHPReader->canRead($fileName)){
            //创建2003版的Excel
            $PHPReader = new PHPExcel_Reader_Excel5();
            if(!$PHPReader->canRead($fileName)){  
                echo "No Excel";  
                return ;  
            }  
        }
        //调用PHPExcel
        $PHPExcel = $PHPReader->load($fileName);
        //读取Excel的第一个工作表
        $currentSheet = $PHPExcel->getSheet(0); 
        $addr = 'A1'; 
        $cell = $currentSheet->getCell($addr)->getValue();
        //将文本转换为字符串
        if($cell instanceof PHPExcel_RichText){
            $cell = $cell->_toString(); 
        }
//        $_status = $cell;注释
        $_status = 'WAIT_SELLER_SEND_GOODS';
//      $_status= ''; //WAIT_SELLER_SEND_GOODS//注释
        $_taobaoConnect= new TaobaoConnector();
        $_taobaoConnect->__url=Yii::app()->params['taobao_api']['url'] ;
        $_taobaoConnect->__appkey= Yii::app()->params['taobao_api']['appkey'] ;
        $_taobaoConnect->__appsecret= Yii::app()->params['taobao_api']['appsecret'] ;
        $_taobaoConnect->__method= Yii::app()->params['taobao_api']['method'] ;
        $_taobaoConnect->__fields= Yii::app()->params['taobao_api']['fields'] ;
        $_orders= $_taobaoConnect->connectTaobao( $_sessionkey,$this->_startDate,$this->_endDate,$_status,$hasNext,$pageNo) ;
        if (array_key_exists('error_response',$_orders)){
            Yii::log('Caught exception: ' . serialize($_orders), 'error', 'system.fail');
            exit(); 
        }
        if (array_key_exists('trades_sold_get_response',$_orders)){
            if (!empty($_orders)){
                return $_orders ;           
            }else{
                Yii::log('No data from'. $this->_startDate . 'to' . $this->_endDate , 'error', 'system.fail');
                exit();
            }
        }else{
            return false;
        }
    }
}

