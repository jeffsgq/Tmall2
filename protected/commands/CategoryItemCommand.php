<?php
Yii::$enableIncludePath = false; 
Yii::import('application.components.TaobaoConnectorItem');
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' ) ;
include_once (dirname(__FILE__).'/../extensions/PHPExcel/PHPExcel/IOFactory.php');
class CategoryItemCommand extends ConsoleCommand{
    protected $PHPExcel = null;
    protected $PHPReader = null;
    protected $PHPWrite = null;
    protected $readFileName = null;
    protected $saveFileName = null;
    protected $_className = null;
    //EXCELTITLE
    protected $categoryTitle = null;
    protected $itemTitle = null;
    protected $category = null;
    public function init(){
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->readFileName = dirname(__FILE__).'/../../Excel/num_iid.xls';
        $this->PHPReader = $this->PHPExcel->load($this->readFileName);
        $this->saveFileName = dirname(__FILE__).'/../../Excel/category.xls';
        $this->PHPWrite = new PHPExcel();
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
        //创建category.xls
        fopen($this->saveFileName, "w+");
        //初始化Excel标题
        $this->categoryTitle = array("num_iid","title","input_str","num","approve_status","cid_1","Category_1","cid_2","Category_2","cid_3","Category_3","cid_4","Category_4");
        $this->itemTitle = array("num_iid","title","input_str","num","approve_status","cid1");
        $this->category = array();
    }
    //执行方法
    public function run($args){
        $this->_Print(); 
    }
     //调整Excel表中属性的顺序
    public function _order(){
        $currentSheet = $this->PHPWrite->setActiveSheetIndex(0);
        $rowN = $currentSheet->getHighestRow();
        $colN = $currentSheet->getHighestColumn();
        $i = ord($colN);
        $currentSheet->setCellValue(chr($i+1).'1' , '叶子类目cid')
                    ->setCellValue(chr($i+2).'1', '叶子类目Category');
        //循环转换
        for($j=2;$j<=$rowN;$j++){
            $flag = $currentSheet->getCell("E".$j)->getValue();
            if(!empty($flag)){
                $index = 70;
                while(!empty($flag)){
                    $flag = $currentSheet->getCell(chr($index).$j)->getValue();
                    $index++;
                }
                $currentSheet->setCellValue(chr($i+1).$j,$currentSheet->getCell(chr($index-3).$j)->getValue())
                    ->setCellValue(chr($i+2).$j,$currentSheet->getCell(chr($index-2).$j)->getValue());
            }
        }
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
        $this->_order();
        $this->_endSaveExcel();
        echo 'END---category.xls';
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
    //出入到Excel
    public function _insertExcel($num_iid,$i){
        //获取API属性
        $_itemsTmallAll = $this->_getAPIValue($num_iid);//一个$num_iid对应一列数据
        if(!empty($_itemsTmallAll)){
            $array = array();
            $array2 = array();
            foreach ($_itemsTmallAll as $_firstKey=>$_firstValue){
                //获取Item数据
                array_push($array, $_firstValue['num_iid'],$_firstValue['title'],$_firstValue['input_str'],$_firstValue['num'],$_firstValue['approve_status']);
                //根据cid获取
                $item = $this->_selectItems($_firstValue['cid']);
                //将c_id存放数据组
                array_push($array2,$_firstValue['cid'],$item['name']);
                $cid = $item['parent_cid'];
                while($cid!=0){ //当$item['parent_cid']为0时结束
                    $item = $this->_selectItems($cid);
                    $cid = $item['parent_cid'];
                    //放入数组中
                    array_push($array2,$item['c_id']);
                    array_push($array2,$item['name']);
                }
            }
            //写入Excel
            for($j=0,$index = 65;$j<count($array);$j++,$index++){
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue(chr($index).$i, $array[$j]);
            }
            //写入Excel
            $this->_orderALine($array2,$i);
            //释放内存
            unset($array);
            unset($array2);
        }else{
            //写入Excel
            $this->PHPWrite->setActiveSheetIndex(0)->setCellValue("A".$i, $num_iid);
        }
    }
    //读取Excel中的数据
    public function _readExcelData($rowIndex,$colIndex){
        $currentSheet = $this->PHPReader->getSheet(0);
        $addr = $colIndex.$rowIndex;
        $cell = $currentSheet->getCell($addr)->getValue();
        return $cell;
    }
     //Excel的头部
    public function _startSaveExcel(){
        $currentSheet = $this->PHPWrite->setactivesheetindex(0);
        for($i=0,$index = 65;$i<count($this->categoryTitle);$i++,$index++){
            $currentSheet->setCellValue(chr($index)."1", $this->categoryTitle[$i]);    
        }
        $currentSheet->setTitle("Sheet1");
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
    //调整一行的数据位置
    public function _orderALine($array,$row){//F开始
        $array = array_reverse($array);
        //插入数据
        $index_j = 70;
        $index_O = 71;
        for($i=0;$i<count($array);$i++){
           if($i%2==0){//偶数
               $this->PHPWrite->setActiveSheetIndex(0)->setCellValue(chr($index_O).$row, $array[$i]);
               $index_O = $index_O + 2;
           }else{
               $this->PHPWrite->setActiveSheetIndex(0)->setCellValue(chr($index_j).$row, $array[$i]);
               $index_j = $index_j + 2;
           }
        }
    }
    //连接淘宝API
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
     //通过cid搜索商品
    public function _selectItems($cid){
        //建立数据库连接
        $connection = Yii::app()->db;
        $item = $connection->createCommand()
                ->select('c_id,is_parent,name,parent_cid')
                ->from('0_parentcid')
                ->where('c_id=:cid',array(':cid'=>$cid))
                ->queryRow();
        return $item;
    }
}
