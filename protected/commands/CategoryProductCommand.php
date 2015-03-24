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
    public function init(){
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->readFileName = dirname(__FILE__).'/../../Excel/num_iid.xls';
        $this->PHPReader = $this->PHPExcel->load($this->readFileName);
        $this->saveFileName = dirname(__FILE__).'/../../Excel/sku.xls';
        $this->PHPWrite = new PHPExcel();
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
    }
    
    public function run($args){
        
        $this->_Print();
//        $_itemsTmallAll = $this->_getAPIValue('2100646557924');
//        print_r($_itemsTmallAll);
    }

    //读取Excel中的数据
    public function _readExcelData($rowIndex,$colIndex){
        
        //读取Excel的第一个工作表
        $currentSheet = $this->PHPReader->getSheet(0);
        //单元格位置
        $addr =$colIndex.$rowIndex;
        $cell = $currentSheet->getCell($addr)->getValue();
        return $cell;
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
        if(!empty($_itemsTmallAll)){//条件1
            foreach ($_itemsTmallAll as $_firstKey=>$_firstValue){
                //存储对应属性
                $num_iid_value = $_firstValue['num_iid'];//1
                $title_value = $_firstValue['title'];//2
                $item_outer_id_value = $_firstValue['outer_id'];//3
                $approve_status_value = $_firstValue['approve_status'];//4
                foreach ($_firstValue as $_secnodKey=>$_secondValue){
                    if(is_array($_secondValue)){//判断数组
                        foreach ($_secondValue as $_thirdKey=>$_thirdValue){//行数   
                           
                            foreach ($_thirdValue as $_fourthKey=>$_fourthValue){//列数
                            //初始化
                                $sku_id_v = null;
                                $sku_outer_id_v = null;
                                $quantity_v = null;
                                $with_hold_quantity_v = null;
                                $price_v = null;
                                $color_v = null;
                                $size_v = null;
                                //获取SKU数据
                                if(array_key_exists("sku_id", $_fourthValue)){
                                    $sku_id_v = $_fourthValue['sku_id'];//5
                                }
                                if(array_key_exists("outer_id", $_fourthValue)){
                                    $sku_outer_id_v = $_fourthValue['outer_id'];//6
                                }
                                if(array_key_exists("quantity", $_fourthValue)){
                                    $quantity_v = $_fourthValue['quantity'];//7
                                }
                                if(array_key_exists("with_hold_quantity", $_fourthValue)){
                                    $with_hold_quantity_v = $_fourthValue['with_hold_quantity'];//8
                                }
                                if(array_key_exists("price", $_fourthValue)){
                                    $price_v = $_fourthValue['price'];//9
                                }
                                //获取颜色和尺码
                                if(array_key_exists("properties_name", $_fourthValue)){
                                    $properties_name_v = $_fourthValue['properties_name'];
                                    $properties_v = $this->_splitStr($properties_name_v);
                                    $color_v = $properties_v[0];//10
                                    $size_v = $properties_v[1];//11
                                }
                                
                                

                                //插入Excel
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A'.$rowIndex,$num_iid_value);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('B'.$rowIndex,$title_value);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('C'.$rowIndex,$item_outer_id_value);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('D'.$rowIndex,$sku_id_v);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('E'.$rowIndex,$sku_outer_id_v);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('F'.$rowIndex,$quantity_v);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('G'.$rowIndex,$with_hold_quantity_v);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('H'.$rowIndex,$price_v);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('I'.$rowIndex,$color_v);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('J'.$rowIndex,$size_v);
                                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('K'.$rowIndex,$approve_status_value);
                                $rowIndex = $rowIndex + 1; 
                            } 
                        }
                    }
                 }
            }
        }else{//条件2
             $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A'.$rowIndex,$num_iid);
             $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('E'.$rowIndex,0);
             $rowIndex = $rowIndex + 1;
        }
        return $rowIndex;
    }
    
     //Excel的头部
    public function _startSaveExcel(){
        
       $this->PHPWrite->setactivesheetindex(0)
            //向Excel中添加数据
            ->setCellValue('A1', 'num_iid')
            ->setCellValue('B1', 'title')
            ->setCellValue('C1', 'item_outer_id')
            ->setCellValue('D1', 'sku_id')
            ->setCellValue('E1', 'sku_outer_id')
            ->setCellValue('F1', 'quantity')
            ->setCellValue('G1', 'with_hold_quantity')
            ->setCellValue('H1', 'price')
            ->setCellValue('I1', 'color')
            ->setCellValue('J1', 'size')
            ->setCellValue('K1', 'approve_status')
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
    
    public function _splitStr($str){
        
        $arr0 = array();
        $arr = explode(':', $str);
        $color = explode(';', $arr[3])[0];
        $size = $arr[6];
        //添加元素到数组
        array_push($arr0,$color, $size);
        return $arr0;
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

