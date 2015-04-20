<?php

Yii::$enableIncludePath = false;
Yii::import('application.components.TaobaoConnector');
Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' );
include_once (dirname(__FILE__) . '/../extensions/PHPExcel/PHPExcel/IOFactory.php');

class InventoryCommand extends ConsoleCommand {

    protected $PHPExcel = null;
    protected $PHPReader = null;
    protected $PHPWrite = null;
    protected $readFileName = null;
    protected $saveFileName = null;
    protected $_className = null;

    public function init() {
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->saveFileName = dirname(__FILE__) . '/../../Excel/Inventory_num_iid.xls';
        $this->PHPWrite = new PHPExcel();
        $this->_className = get_class();
        $this->beforeAction($this->_className, '');
        fopen($this->saveFileName, "w+");
    }

    public function run($args) {
        $this->_Print();
    }

    //获取API属性
    public function _getAPIValue($page_no, $page_size, $banner) {
        //num_iid不存在则返回NULL
        $_itemsTmallAll = array();
        $_itemsTmall = $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'], $page_no, $page_size, $banner);
        if (!empty($_itemsTmall)) {
            if (array_key_exists('items', $_itemsTmall['items_inventory_get_response'])) {  //['items']
                array_push($_itemsTmallAll, $_itemsTmall['items_inventory_get_response']['items']['item']);
            } else {
                array_push($_itemsTmallAll, null);
            }
            return $_itemsTmallAll;
        } else {

            return $_itemsTmall;
        }
    }

    private function _connectTmall($_sessionkey, $page_no, $page_size, $banner) {

        $_taobaoConnect = new TaobaoConnector();
        $_taobaoConnect->__url = Yii::app()->params['taobao_api']['url'];
        $_taobaoConnect->__appkey = Yii::app()->params['taobao_api']['appkey'];
        $_taobaoConnect->__appsecret = Yii::app()->params['taobao_api']['appsecret'];
        $_taobaoConnect->__method = Yii::app()->params['taobao_api']['methods']['commodity_inventory_method'];
        $_taobaoConnect->__fields = Yii::app()->params['taobao_api']['fields']['commodity_onsale_inventory_field'];
        $_items = $_taobaoConnect->connectTaobaoinventory($_sessionkey, $page_no, $page_size, $banner);
        if (array_key_exists('error_response', $_items)) {
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
            return NULL;
        }
        if (array_key_exists('items_inventory_get_response', $_items)) {
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
                ->setCellValue('A1', 'Inventory_num_iid')
                ->setCellValue('B1', 'Banner')
                ->setTitle('sheet1');
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

    public function _Print1st($page_no, $page_size, $rowIndex) {
        $banner = 'for_shelved';
        $_total_results = $this->_getTotalResults($page_no, $page_size, $banner);
        $_page = floor($_total_results / $page_size) + 1;
        do {
            $_itemsTmallAll = $this->_getAPIValue($page_no, $page_size, $banner);
            foreach ($_itemsTmallAll as $_firstKey => $_firstValue) {
                foreach ($_firstValue as $_secnodKey => $_secondValue) {
                    print_r($_secondValue);
                    $_inventory_num_iid = null;
                    if (array_key_exists("num_iid", $_secondValue)) {
                        $_inventory_num_iid = $_secondValue['num_iid'];
                    }
                    //插入Excel	
                    $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A' . $rowIndex, $_inventory_num_iid);
                    $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('B' . $rowIndex, $banner);
                    $rowIndex = $rowIndex + 1;
                }
                $page_no = $page_no + 1;
                $_page = $_page - 1;
            }
        } while (!$_page == 0); 
        echo '---for_shelved---';
        return $rowIndex;
    }

    public function _Print2st($page_no, $page_size, $row) {
        $banner = 'sold_out';  //never_on_shelf
        $_total_results = $this->_getTotalResults($page_no, $page_size, $banner);
        $_page = floor($_total_results / $page_size) + 1;
        do {
            $_itemsTmallAll = $this->_getAPIValue($page_no, $page_size, $banner);

            if (current($_itemsTmallAll) == FALSE) {
//                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A' . $row, "0000");
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('B' . $row, $banner);
                break;
            } else {
                foreach ($_itemsTmallAll as $_firstKey => $_firstValue) {
                    foreach ($_firstValue as $_secnodKey => $_secondValue) {
                        print_r($_secondValue);
                        $_inventory_num_iid = null;
                        if (array_key_exists("num_iid", $_secondValue)) {
                            $_inventory_num_iid = $_secondValue['num_iid'];
                        }
                        //插入Excel	
                        $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A' . $row, $_inventory_num_iid);
                        $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('B' . $row, $banner);
                        $row = $row + 1;
                    }
                    $page_no = $page_no + 1;
                    $_page = $_page - 1;
                }
            }
        } while (!$_page == 0);

        echo '---sold_out---';
    }

    //循环输出并保存在Excel中
    public function _Print() {
        ob_start();
        $this->_startSaveExcel();
        $page_no = 1;
        $page_size = 200;
        $rowIndex = 2;
        $row = $this->_Print1st($page_no, $page_size, $rowIndex);
        $this->_Print2st($page_no, $page_size, $row);
        $this->_endSaveExcel(); //Excel的尾部
        echo '--END--';
    }

    public function _getTotalResults($page_no, $page_size, $banner) {
        $_itemsTmallNoAll = array();
        $_itemsTmallNo = $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'], $page_no, $page_size, $banner);
        if (!empty($_itemsTmallNo)) {
            if (array_key_exists('total_results', $_itemsTmallNo['items_inventory_get_response'])) {
                array_push($_itemsTmallNoAll, $_itemsTmallNo['items_inventory_get_response']['total_results']);
            }

            return $_itemsTmallNoAll[0];
        } else {
            return $_itemsTmallNo;
        }
    }

}