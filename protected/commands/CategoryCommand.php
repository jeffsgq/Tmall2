<?php

Yii::$enableIncludePath = false; 
Yii::import('application.components.TaobaoConnectorItem');
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' );
include_once (dirname(__FILE__).'/../extensions/PHPExcel/PHPExcel/IOFactory.php');

class CategoryCommand extends ConsoleCommand {
    
    protected $PHPExcel = null;
    protected $PHPReader = null;
    protected $PHPWrite = null;
    protected $readFileName = null;
    protected $saveFileName = null;
    protected $_className= null ;
    public function init(){
        
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->readFileName = dirname(__FILE__).'/../../Excel/num_iid.xls';
        $this->PHPReader = $this->PHPExcel->load($this->readFileName);
        $this->saveFileName = dirname(__FILE__).'/../../Excel/item.xls';
        $this->PHPWrite = new PHPExcel();
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
    }
    
    public function run($args){
        
//        $this->_Print();
    }

    //读取Excel中的数据
    public function _readExcelData($rowIndex,$colIndex){
        
        $currentSheet = $this->PHPReader->getSheet(0);
        $addr =$colIndex.$rowIndex;
        $cell = $currentSheet->getCell($addr)->getValue();
        return $cell;
    }
    
    //获取API属性
    public function _getAPIValue($num_iid){
        
        $_itemsTmallAll= array();
        $_itemsTmall= $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'],$num_iid."");
        if(!empty($_itemsTmall)){
            if (array_key_exists('item',$_itemsTmall['item_get_response'])){
                array_push($_itemsTmallAll, $_itemsTmall['item_get_response']['item']);
            }
            return $_itemsTmallAll;
        }else{
            return $_itemsTmall;
        }
    }
    
    public function _insertExcel($num_iid,$i){
        
        //获取API属性
        $_itemsTmallAll = $this->_getAPIValue($num_iid);//一个$num_iid对应一列数据
        if(!empty($_itemsTmallAll)){
            foreach ($_itemsTmallAll as $_firstKey=>$_firstValue){
                //获取Item数据
                $num_iid_value = $_firstValue['num_iid'];//1
                $title_value = $_firstValue['title'];//2
                $input_str_value = $_firstValue['input_str'];//3
                $num_value = $_firstValue['num'];//4
                $approve_status_value = $_firstValue['approve_status'];//5
                $cid_value = $_firstValue['cid'];//6

                //插入Excel
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A'.$i, $num_iid_value);
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('B'.$i, $title_value);
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('C'.$i, $input_str_value);
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('D'.$i, $num_value);
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('E'.$i, $approve_status_value);
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('F'.$i, $cid_value);
            }
        }else{
            $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A'.$i, $num_iid);
        }
    }
    
    //Excel的头部
    public function _startSaveExcel(){
        
       $this->PHPWrite->setactivesheetindex(0)
            //向Excel中添加数据
            ->setCellValue('A1', 'num_iid')
            ->setCellValue('B1', 'title')
            ->setCellValue('C1', 'input_str')
            ->setCellValue('D1', 'num')
            ->setCellValue('E1', 'approve_status')
            ->setCellValue('F1', 'cid1')
            ->setTitle('sheet1');
    }
    
    //Excel的尾部
    public function _endSaveExcel(){
        
        if(!is_writable($this->saveFileName)){
            echo 'Can not Write';
            exit();
        }
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename='.$this->saveFileName);
        header('Cache-Control: max-age=0');
        //创建文件使用Excel2003版本
        $objWriter = PHPExcel_IOFactory::createWriter($this->PHPWrite,'Excel5');  
        $objWriter->save($this->saveFileName);
    }
    
    //循环输出并保存在Excel中
    public function _Print(){
        
        ob_start();
        $this->_startSaveExcel();
        $currentSheet = $this->PHPReader->getSheet(0);
        $allRow = $currentSheet->getHighestRow();
        //循环写入
        for($rowIndex=2;$rowIndex<=$allRow;$rowIndex++){
            $num_iid = $this->_readExcelData($rowIndex, 'A');
            $this->_insertExcel($num_iid,$rowIndex);
        }
        $this->_endSaveExcel();
        echo 'END';
    }
    
    private function _connectTmall($_sessionkey,$num_iid){
        
        $_taobaoConnect= new TaobaoConnectorItem();
        $_taobaoConnect->__url=Yii::app()->params['taobao_api']['url'] ;
        $_taobaoConnect->__appkey= Yii::app()->params['taobao_api']['appkey'] ;
        $_taobaoConnect->__appsecret= Yii::app()->params['taobao_api']['appsecret'] ;
        $_taobaoConnect->__method= Yii::app()->params['taobao_api']['method2'] ;
        $_taobaoConnect->__fields= Yii::app()->params['taobao_api']['fields2'] ;
        $_items= $_taobaoConnect->connectTaobaoItem( $_sessionkey,$num_iid) ;
        if (array_key_exists('error_response',$_items)){
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
//            exit(); 
            return NULL;
        }
        if (array_key_exists('item_get_response',$_items)){
            if (!empty($_items)){
                return $_items ;           
            }else{
                Yii::log('No data parent_cid'.$num_iid, 'error', 'system.fail');
//                exit();
                return NULL;
            }
        }else{
            return NULL;
        }
    }
}

