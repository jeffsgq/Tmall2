<?php
Yii::$enableIncludePath = false;
Yii::import('application.components.TaobaoConnectorSKU') ;
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' ) ;
include_once (dirname(__FILE__).'/../extensions/PHPExcel/PHPExcel/IOFactory.php');
PHPExcel_CachedObjectStorageFactory::cache_in_memory_serialized;
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
    protected $_parentFields= array();
    protected $_skuFields= array();

    public function init(){
        ini_set('memory_limit', '800M');
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->PHPWrite = new PHPExcel();
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
        //数组初始化
        $this->_parentFields = array("num_iid","banner","title","outer_id","approve_status","skus");//skus must be the last one
        $this->_skuFields = array("sku_id","outer_id","quantity","with_hold_quantity","price","properties_name");
        $this->titleArray = array("num_iid","banner","title","item_outer_id","approve_status","sku_id","sku_outer_id","quantity","with_hold_quantity","price","properties");
    }
    public function run($choic){
        $this->_prompt($choic);
        switch ($choic[0]) {
            case 'inventory':
                $this->readFileName = dirname(__FILE__).'/../../Excel/Inventory_num_iid.xls';
                $this->saveFileName = dirname(__FILE__).'/../../Excel/Inventory_sku.xls';
                break;
            case 'onsale':
                $this->readFileName = dirname(__FILE__).'/../../Excel/Onsale_num_iid.xls';
                $this->saveFileName = dirname(__FILE__).'/../../Excel/Onsale_sku.xls';
                break;
            default:
                echo "**************************************************\n"
                    . "Please input parameter : onsale or inventory\n"
                    . "**************************************************";
                exit();
        }
        $this->PHPReader = $this->PHPExcel->load($this->readFileName);
        fopen($this->saveFileName, "w+");
        $this->_generateExcel();
    }
    //parameter prompt
    public function _prompt($args){
        if(empty($args)){
            echo "**************************************************\n"
            . "Please input parement : onsale or inventory\n"
            . "Like that:categoryproduct onsale\n"
            . "**************************************************";
            exit();
        }
    }
    //获取API属性
    public function _getAPIValue($num_iid){
        //num_iid不存在则返回NULL
        $_itemsTmallAll= array();
        $_itemsTmall= $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'],$num_iid);
        if(!empty($_itemsTmall)){
            if (array_key_exists('item',$_itemsTmall['item_get_response'])){
                $_itemsTmallAll= $_itemsTmall['item_get_response']['item'];
            }
            return $_itemsTmallAll;
        }else{
            return $_itemsTmall;
        }
    }
    
    public function _filterApiParentValue($num_iid,$banner){
        $_filterResult= array();
        $_itemsTmallAll = $this->_getAPIValue($num_iid);
//        print_r($_itemsTmallAll);
        if(!empty($_itemsTmallAll)){
            foreach ($this->_parentFields as $field){
                $_filterResult['banner']= $banner;
                if(array_key_exists($field,$_itemsTmallAll)){
//                    array_push($_filterResult,$_itemsTmallAll[$fields]);
                    $_filterResult[$field]= $_itemsTmallAll[$field];
                }else{
                    $_filterResult[$field]= "";
                }
            }
            unset($_itemsTmallAll);
            return $_filterResult;
        }else{
            Yii::log('Caught exception: num_iid:' .$num_iid. 'item not exists', 'error', 'system.fail');
            return false;
        }
    }
    
    public function _filterApiSkuValue($_filterParentResult,$num_iid){
        if(!empty($_filterParentResult['skus'])){
            if(array_key_exists('sku', $_filterParentResult['skus'])){
                foreach ($_filterParentResult['skus']['sku'] as $key=> $value){
                    foreach($this->_skuFields as $field){
                        if(!array_key_exists($field, $value)){
                            $_filterParentResult['skus']['sku'][$key][$field]= "";
                        }
                    }
                }
            }else{
                Yii::log('Caught exception: num_iid:' .$num_iid. 'sku not exists', 'error', 'system.fail');
            }
        }else{
            $_filterParentResult['skus']['sku']= array();
            Yii::log('Caught exception: num_iid:' .$num_iid. 'skus not exists', 'error', 'system.fail');
        }
        return $_filterParentResult;
    }
    
    public function _insertExc($_numID,$banner,$_row){
        $currentSheet = $this->PHPWrite->setActiveSheetIndex(0);
        $_filterParentResult= $this->_filterApiParentValue($_numID,$banner);
        $_filterResult = $this->_filterApiSkuValue($_filterParentResult, $_numID);
//        print_r($_filterResult);exit();
        unset($_filterParentResult);
        if(count($_filterResult['skus']['sku'])==0){
            $_skuQty=1;
        }else{
            $_skuQty= count($_filterResult['skus']['sku']);
        }

        for($i=0;$i<$_skuQty;$i++){
            $index=65;
            foreach ($this->_parentFields as $field){
                if(!is_array($_filterResult[$field])){
                    $currentSheet->setCellValue(chr($index).($_row+$i),$_filterResult[$field]);
                    $index++;
                }
            }
            foreach ($this->_skuFields as $sku) {
                if (empty($_filterResult['skus']['sku'])){
                    $currentSheet->setCellValue(chr($index).($_row+$i),NULL);
                }else{
                    $currentSheet->setCellValue(chr($index).($_row+$i),$_filterResult['skus']['sku'][$i][$sku]);
                }
                $index++;
            }
        }
        $_newRow= $_row + $_skuQty;
        return $_newRow;
    }
    
    public function _generateExcel(){
        ob_start();
        $this->_startSaveExcel();//Excel的头部
        $currentSheet = $this->PHPReader->getSheet(0);
        $allRow = $currentSheet->getHighestRow();
        //循环写入
        $rowIndex = 2;
        for($rowI=2;$rowI<=$allRow;$rowI++){
            $num_iid = $this->_readExcelData($rowI, 'A');
            $banner = $this->_readExcelData($rowI, 'B');
            if(!empty($num_iid)){
               $rowIndex = $this->_insertExc($num_iid,$banner,$rowIndex); 
            }
//            echo $rowIndex;
        }
        $this->_endSaveExcel();//Excel的尾部
        echo '-------------END-------------';
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
                return $_items ;           
        }else{
            Yii::log('item_get_response not exists:'.$num_iid, 'error', 'system.fail');
            return NULL;
        }
    }
}

