<?php
Yii::$enableIncludePath = false;
Yii::import('application.components.TaobaoConnectorSKU') ;
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' ) ;
include_once (dirname(__FILE__).'/../extensions/PHPExcel/PHPExcel/IOFactory.php');
class CategoryProductCommand extends ConsoleCommand { 
    protected $PHPExcel = null;
    protected $PHPReader = null;
    protected $PHPWrite = null;
    protected $readFileName = null;
    protected $saveFileName = null;
    protected $_className= null ;
    //SKU、ITEM、EXCELTITLE的数组
    protected $skuArray = null;
    protected $itemArray = null;
    protected $titleArray = null;
    public function init(){
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->readFileName = dirname(__FILE__).'/../../Excel/num_iid.xls';
        $this->PHPReader = $this->PHPExcel->load($this->readFileName);
        $this->saveFileName = dirname(__FILE__).'/../../Excel/sku.xls';
        $this->PHPWrite = new PHPExcel();
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
        //创建sku.xls
        fopen($this->saveFileName, "w+");
        //数组初始化
        $this->skuArray = array("sku_id","outer_id","quantity","with_hold_quantity","price","properties_name");
        $this->itemArray = array("num_iid","title","outer_id","approve_status");
        $this->titleArray = array("num_iid","title","item_outer_id","approve_status","sku_id","sku_outer_id","quantity","with_hold_quantity","price","properties");
    }
    public function run($args){
        $this->_Print();
    }
    //获取API属性
    public function _getAPIValue($num_iid){
        //num_iid不存在则返回NULL
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
    //写入Excel
    public function _insertExcel($num_iid,$rowIndex){
        //获取API属性
        $_itemsTmallAll = $this->_getAPIValue($num_iid);
        $currentSheet = $this->PHPWrite->setActiveSheetIndex(0);
        if(!empty($_itemsTmallAll)){//条件1
            foreach ($_itemsTmallAll as $_firstKey=>$_firstValue){
                //初始化数组ITEM
                $item_Array = array("num_iid"=>null,"title"=>null,"outer_id"=>null,"approve_status"=>null);
                //打印出TITLE
                $flag = false;
                $_secondVA = null;
                foreach ($_firstValue as $_secnodKey=>$_secondValue){
                    //存储对应属性
                    foreach ($item_Array as $key => $value) {
                        if($_secnodKey==$key){
                            $item_Array[$key] = $_firstValue[$key];
                        }
                    }
                    if(is_array($_secondValue)){
                        $_secondVA = $_secondValue;
                        $flag = true;
                    }
                 }
                if($flag){
                    foreach ($_secondVA as $_thirdKey=>$_thirdValue){//行数 
                            foreach ($_thirdValue as $_fourthKey=>$_fourthValue){//列数
                                //初始化数组SKU
                                $sku_Array = array("sku_id"=>null,"outer_id"=>null,"quantity"=>null,"with_hold_quantity"=>null,"price"=>null,"properties_name"=>null);
                                foreach ($_fourthValue as $_fifthhKey => $_fifthhValue) {
                                    //存储对应属性
                                    foreach ($sku_Array as $key => $value) {
                                        if($_fifthhKey==$key){
                                            $sku_Array[$key] = $_fifthhValue;
                                        }
                                    }   
                                }
                                //将数据插入到EXCEL中
                                $index = 65;
                                foreach ($item_Array as $key => $value) {
                                    $currentSheet->setCellValue(chr($index).$rowIndex,$item_Array[$key]);
                                    $index ++;
                                }
                                foreach ($sku_Array as $key => $value) {
                                    $currentSheet->setCellValue(chr($index).$rowIndex,$sku_Array[$key]);
                                    $index ++;
                                }
                                $rowIndex = $rowIndex + 1; 
                            } 
                        }
                }
            }
        }else{//条件2
             $currentSheet->setCellValue('A'.$rowIndex,$num_iid);
             $rowIndex ++;
        }
        return $rowIndex;
    }
    //读取Excel中的数据
    public function _readExcelData($rowIndex,$colIndex){
        //单元格位置
        $addr =$colIndex.$rowIndex;
        $cell = $this->PHPReader->setactivesheetindex(0)->getCell($addr)->getValue();
        return $cell;
    }
    //Excel的头部
    public function _startSaveExcel(){
        $currentSheet = $this->PHPWrite->setactivesheetindex(0);
        for($i=0,$index = 65;$i<count($this->titleArray);$i++,$index++){
            $currentSheet->setCellValue(chr($index)."1", $this->titleArray[$i]);    
        }
        $this->PHPWrite->setactivesheetindex(0)->setTitle("Sheet1");
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
        $this->_startSaveExcel();//Excel的头部
        $currentSheet = $this->PHPReader->getSheet(0);
        $allRow = $currentSheet->getHighestRow();
        //循环写入
        $rowIndex = 2;
        for($rowI=2;$rowI<=$allRow;$rowI++){
            $num_iid = $this->_readExcelData($rowI, 'A');
            $rowIndex = $this->_insertExcel($num_iid,$rowIndex);
        }
        $this->_endSaveExcel();//Excel的尾部
        echo 'END--sku.xml';
    }
    private function _connectTmall($_sessionkey,$num_iid){
        $_taobaoConnect= new TaobaoConnectorSKU();
        $_taobaoConnect->__url=Yii::app()->params['taobao_api']['url'] ;
        $_taobaoConnect->__appkey= Yii::app()->params['taobao_api']['appkey'] ;
        $_taobaoConnect->__appsecret= Yii::app()->params['taobao_api']['appsecret'] ;
        $_taobaoConnect->__method= Yii::app()->params['taobao_api']['method3'] ;
        $_taobaoConnect->__fields= Yii::app()->params['taobao_api']['fields3'] ;
        $_items= $_taobaoConnect->connectTaobaoSKU( $_sessionkey,$num_iid) ;
        if (array_key_exists('error_response',$_items)){
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
            return NULL;
        }
        if (array_key_exists('item_get_response',$_items)){
            if (!empty($_items)){
                return $_items ;           
            }else{
                Yii::log('No data parent_cid'.$num_iid, 'error', 'system.fail');
                return NULL;
            }
        }else{
            return NULL;
        }
    }
}

