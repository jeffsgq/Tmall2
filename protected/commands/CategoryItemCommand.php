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
    //Category
    protected $PHPReader2 = null;
    protected $PHPWrite2 = null;
    protected $readFileName2 = null;
    protected $saveFileName2 = null;

    public function init(){

        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->readFileName = dirname(__FILE__).'/../../Excel/item.xls';
        $this->PHPReader = $this->PHPExcel->load($this->readFileName);
        $this->saveFileName = dirname(__FILE__).'/../../Excel/category.xls';
        $this->PHPWrite = new PHPExcel();
        //Category
        $this->readFileName2 = dirname(__FILE__).'/../../Excel/num_iid.xls';
        $this->PHPReader2 = $this->PHPExcel->load($this->readFileName2);
        $this->saveFileName2 = dirname(__FILE__).'/../../Excel/item.xls';
        $this->PHPWrite2 = new PHPExcel();
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
    }
    
    //执行方法
    public function run($args){
        
        $this->_Print2();
        $this->_Print();
        
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
            ->setCellValue('F1', 'cid_1')
            ->setCellValue('G1', 'Category_1')
            ->setCellValue('H1', 'cid_2')
            ->setCellValue('I1', 'Category_2')
            ->setCellValue('J1', 'cid_3')
            ->setCellValue('K1', 'Category_3')
            ->setCellValue('L1', 'cid_4')
            ->setCellValue('M1', 'Category_4')
            ->setCellValue('N1', 'cid_5')
            ->setCellValue('O1', 'Category_5')
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

    
    //写入数据
    public function _Print(){
        
        ob_start();
        $this->_startSaveExcel();
        
        $currentSheet = $this->PHPReader->getSheet(0);
        $rowN = $currentSheet->getHighestRow();
        $colIndex = 'F';
        //循环读取Excel中的数据
        for($rowIndex = 2;$rowIndex <= $rowN;$rowIndex++){
            $cell = $this->_readExcelData($rowIndex,$colIndex); //$cell为cid
            if(!empty($cell)){
                $item = $this->_selectItems($cell); //$item为当前cid的对应值
                $cid = $item['parent_cid'];
                //Excel初始化
                $this->_writeExcelData($this->_readExcelData($rowIndex,'A'), $rowIndex,'A');
                $this->_writeExcelData($this->_readExcelData($rowIndex,'B'), $rowIndex,'B');
                $this->_writeExcelData($this->_readExcelData($rowIndex,'C'), $rowIndex,'C');
                $this->_writeExcelData($this->_readExcelData($rowIndex,'D'), $rowIndex,'D');
                $this->_writeExcelData($this->_readExcelData($rowIndex,'E'), $rowIndex,'E');
                $j = 72;//H
                $this->_writeExcelData($item['c_id'], $rowIndex,'F');
                $this->_writeExcelData($item['name'], $rowIndex,'G');
                while($cid!=0){//当$item['parent_cid']为0时结束
                    $item = $this->_selectItems($cid);
                    $cid = $item['parent_cid'];
                    //数据写入到Excel中
                    $this->_writeExcelData($item['c_id'], $rowIndex, chr($j));
                    $j = $j + 1;
                    $this->_writeExcelData($item['name'], $rowIndex, chr($j));
                    $j = $j + 1;
                }
            }else{
                $this->_writeExcelData($this->_readExcelData($rowIndex,'A'), $rowIndex,'A');
            }
        }
        $this->_order();
        $this->_endSaveExcel();
        echo 'END2--item.xml';
    }
    
    //调整Excel表中属性的顺序
    public function _order(){
        
        $currentSheet = $this->PHPWrite->getSheet(0);
        $rowN = $currentSheet->getHighestRow();
        $colN = $currentSheet->getHighestColumn();
        $i = ord($colN );
        $currentSheet->setCellValue(chr($i+1).'1' , '叶子类目cid')
                    ->setCellValue(chr($i+2).'1', '叶子类目Category');
        //循环转换
        for($j=2;$j<=$rowN;$j++){
            $cid = $currentSheet->getCell('F'.$j)->getValue();
            $Category = $currentSheet->getCell('G'.$j)->getValue();
            //转换
            $this->_orderALine($j, $i);
            //插入
             $currentSheet->setCellValue(chr($i+1).$j , $cid)
                    ->setCellValue(chr($i+2).$j, $Category);  
        }
    }
    
    //调整一行的数据位置
    public function _orderALine($row,$colN){//F开始
        
        $array = array();
        $index = 0;
        $curSheet = $this->PHPWrite->getSheet(0);
        //获取数据
        for($i=70;$i<$colN;$i++){
            if(!empty($curSheet->getCell(chr($i).$row)->getValue())){
                $array[$index] = $curSheet->getCell(chr($i).$row)->getValue();
                $index = $index + 1;
            }else{
                //如果为空则退出本次循环
                break;
            }
        }
        $array = array_reverse($array);
        //插入数据
        $index_j = 70;
        $index_O = 71;
        for($i=0;$i<count($array);$i++){
           if($i%2==0){//偶数
               $curSheet->setCellValue(chr($index_O).$row, $array[$i]);
               $index_O = $index_O + 2;
           }else{
               $curSheet->setCellValue(chr($index_j).$row, $array[$i]);
               $index_j = $index_j + 2;
           }
        }
    }
    
    //将数据写入到Excel中
    public function _writeExcelData($data,$rowIndex,$colIndex){
        $addr = $colIndex.$rowIndex;
        $this->PHPWrite->setActiveSheetIndex(0)->setCellValue($addr,$data);
    }
    
    //读取Excel中的数据
    public function _readExcelData($rowIndex,$colIndex){
        $currentSheet = $this->PHPReader->getSheet(0);
        $addr = $colIndex.$rowIndex;
        $cell = $currentSheet->getCell($addr)->getValue();
        return $cell;
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
    
    //Category
    //读取Excel中的数据
    public function _readExcelData2($rowIndex,$colIndex){
        
        $currentSheet = $this->PHPReader2->getSheet(0);
        $addr =$colIndex.$rowIndex;
        $cell = $currentSheet->getCell($addr)->getValue();
        return $cell;
    }
    //获取API属性
    public function _getAPIValue2($num_iid){
        
        $_itemsTmallAll= array();
        $_itemsTmall= $this->_connectTmall2(Yii::app()->params['taobao_api']['accessToken'],$num_iid."");
        if(!empty($_itemsTmall)){
            if (array_key_exists('item',$_itemsTmall['item_get_response'])){
                array_push($_itemsTmallAll, $_itemsTmall['item_get_response']['item']);
            }
            return $_itemsTmallAll;
        }else{
            return $_itemsTmall;
        }
    }
     public function _insertExcel2($num_iid,$i){
        
        //获取API属性
        $_itemsTmallAll = $this->_getAPIValue2($num_iid);//一个$num_iid对应一列数据
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
                $this->PHPWrite2->setActiveSheetIndex(0)->setCellValue('A'.$i, $num_iid_value);
                $this->PHPWrite2->setActiveSheetIndex(0)->setCellValue('B'.$i, $title_value);
                $this->PHPWrite2->setActiveSheetIndex(0)->setCellValue('C'.$i, $input_str_value);
                $this->PHPWrite2->setActiveSheetIndex(0)->setCellValue('D'.$i, $num_value);
                $this->PHPWrite2->setActiveSheetIndex(0)->setCellValue('E'.$i, $approve_status_value);
                $this->PHPWrite2->setActiveSheetIndex(0)->setCellValue('F'.$i, $cid_value);
            }
        }else{
            $this->PHPWrite2->setActiveSheetIndex(0)->setCellValue('A'.$i, $num_iid);
        }
    }
     //Excel的头部
    public function _startSaveExcel2(){
        
       $this->PHPWrite2->setactivesheetindex(0)
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
    public function _endSaveExcel2(){
        
        if(!is_writable($this->saveFileName)){
            echo 'Can not Write';
            exit();
        }
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename='.$this->saveFileName2);
        header('Cache-Control: max-age=0');
        //创建文件使用Excel2003版本
        $objWriter = PHPExcel_IOFactory::createWriter($this->PHPWrite2,'Excel5');  
        $objWriter->save($this->saveFileName2);
    }
    //循环输出并保存在Excel中
    public function _Print2(){
        
        ob_start();
        $this->_startSaveExcel2();
        $currentSheet = $this->PHPReader2->getSheet(0);
        $allRow = $currentSheet->getHighestRow();
        //循环写入
        for($rowIndex=2;$rowIndex<=$allRow;$rowIndex++){
            $num_iid = $this->_readExcelData2($rowIndex, 'A');
            $this->_insertExcel2($num_iid,$rowIndex);
        }
        $this->_endSaveExcel2();
        echo 'END1'."\n";
    }
     private function _connectTmall2($_sessionkey,$num_iid){
        
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
