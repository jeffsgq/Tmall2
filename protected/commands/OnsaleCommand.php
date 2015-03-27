<?php

Yii::$enableIncludePath = false;
Yii::import('application.components.TaobaoConnectorOnSale');
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' );
include_once (dirname(__FILE__) . '/../extensions/PHPExcel/PHPExcel/IOFactory.php');

class OnsaleCommand extends ConsoleCommand {

    protected $PHPExcel = null;
    protected $PHPReader = null;
    protected $PHPWrite = null;
    protected $readFileName = null;
    protected $saveFileName = null;
    protected $_className = null;
    //OnSale 数组
    protected $onsaleArray = null;

    public function init() {
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
//        $this->readFileName = dirname(__FILE__).'/../../Excel/num_iid.xls';
//        $this->PHPReader = $this->PHPExcel->load($this->readFileName);
        $this->saveFileName = dirname(__FILE__) . '/../../Excel/OnSale_num_iid.xls';
        $this->PHPWrite = new PHPExcel();
        $this->_className = get_class();
        $this->beforeAction($this->_className, '');
//        fopen($this->saveFileName, "w+");
//        //数组初始化
//        $this->onsaleArray = array();
    }

    public function run($args) {
        $this->_Print();
    }

    //获取API属性
    public function _getAPIValue() {
        //num_iid不存在则返回NULL
        $_itemsTmallAll = array();
        $_itemsTmall = $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken']);
        if (!empty($_itemsTmall)) {
            if (array_key_exists('item', $_itemsTmall['items_onsale_get_response']['items'])) {
                array_push($_itemsTmallAll, $_itemsTmall['items_onsale_get_response']['items']['item']);
            }
            return $_itemsTmallAll;
        } else {
            return $_itemsTmall;
        }
    }

    private function _connectTmall($_sessionkey) {

        $_taobaoConnect = new TaobaoConnectorOnSale();
        $_taobaoConnect->__url = Yii::app()->params['taobao_api']['url'];
        $_taobaoConnect->__appkey = Yii::app()->params['taobao_api']['appkey'];
        $_taobaoConnect->__appsecret = Yii::app()->params['taobao_api']['appsecret'];
        $_taobaoConnect->__method = Yii::app()->params['taobao_api']['method4'];
        $_taobaoConnect->__fields = Yii::app()->params['taobao_api']['fields4'];
        $_items = $_taobaoConnect->connectTaobaoonsale($_sessionkey);


        if (array_key_exists('error_response', $_items)) {
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
//            exit(); 
            return NULL;
        }
        if (array_key_exists('items_onsale_get_response', $_items)) {
            if (!empty($_items)) {
                return $_items;
            } else {
                Yii::log('No data numiid', 'error', 'system.fail');
//                exit();
                return NULL;
            }
        } else {
            return NULL;
        }
    }

    //Excel的头部
    public function _startSaveExcel() {

        $this->PHPWrite->setactivesheetindex(0)
                //向Excel中添加数据
                ->setCellValue('A1', 'OnSale_num_iid')
                ->setTitle('sheet1');
    }

    //写入Excel
    public function _insertExcel($rowIndex) {
        //获取API属性
        $_itemsTmallAll = $this->_getAPIValue();
        foreach ($_itemsTmallAll as $_firstKey => $_firstValue) {
            foreach ($_firstValue as $_secnodKey => $_secondValue) {
                print_r($_secondValue);
         
                $_onsale_num_iid = null;
                if (array_key_exists("num_iid", $_secondValue)){
                       $_onsale_num_iid = $_secondValue['num_iid'];
                }
                //插入Excel
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A' . $rowIndex, $_onsale_num_iid);
                $rowIndex = $rowIndex + 1;
            }
        }
        return $rowIndex;
    }

    //Excel的尾部
    public function _endSaveExcel() {

        if (!is_writable($this->saveFileName)) {
            echo 'Can not Write';
            exit();
        }
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename=' . $this->saveFileName);
        header('Cache-Control: max-age=0');
        //创建文件使用Excel2003版本
        $objWriter = PHPExcel_IOFactory::createWriter($this->PHPWrite, 'Excel5');
        $objWriter->save($this->saveFileName);
    }

    //循环输出并保存在Excel中
    public function _Print() {

        ob_start();
        $this->_startSaveExcel();
        //Excel的头部
        //$currentSheet = $this->PHPReader->getSheet(0);
        $rowIndex = 2;
        //循环写入
        $this->_insertExcel($rowIndex);

        $this->_endSaveExcel(); //Excel的尾部
        echo 'END--sku.xml';
    }

    public function _splitStr($str) {

        $arr0 = array();
        $arr = explode(':', $str);
        $color = explode(';', $arr[3])[0];
        $size = $arr[6];
        //添加元素到数组
        array_push($arr0, $color, $size);
        return $arr0;
    }

    public function _formatArray($_orders) {
        $_ordersArray = array();
        foreach ($_orders as $_firstKey => $_firstValue) {
//            foreach ($_firstValue as $_secnodKey=> $_secondValue){
            $_ordersArray[] = $_firstValue;
//            }
        }
        return $_ordersArray;
    }

}
