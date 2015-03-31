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
    
      public function init() {
        $this->PHPExcel = new PHPExcel_Reader_Excel5();
        $this->saveFileName = dirname(__FILE__) . '/../../Excel/OnSale_num_iid.xls';
        $this->PHPWrite = new PHPExcel();
        $this->_className = get_class();
        $this->beforeAction($this->_className, '');
    }
    
      public function run($args) {
            $this->_print();
    }
       
    public function _print(){
        ob_start();
        $this->_startSaveExcel();
        $page_no = 1;
        $page_size = 200;
        $_total_results = $this->_getTotalResults($page_no,$page_size);
        $_page = floor($_total_results / $page_size ) + 1;
        $rowIndex =2;
        do {
            $_itemsTmallAll = $this->_getAPIValue($page_no,$page_size);
            foreach ($_itemsTmallAll as $_firstKey => $_firstValue) 
                {
            foreach ($_firstValue as $_secnodKey => $_secondValue)
                {
            print_r($_secondValue);
             $_onsale_num_iid = null;
                if (array_key_exists("num_iid", $_secondValue)) {
                    $_onsale_num_iid = $_secondValue['num_iid'];
                }
                //插入Excel	
                $this->PHPWrite->setActiveSheetIndex(0)->setCellValue('A' . $rowIndex, $_onsale_num_iid);
                $rowIndex = $rowIndex + 1;
                }
            $page_no = $page_no + 1;
            $_page = $_page - 1;
           
                } 
           }while(!$_page==0);
         
           $this->_endSaveExcel();  
        //Excel的尾部
        echo '--END--';
    }

    //Excel的头部
    public function _startSaveExcel() {

        $this->PHPWrite->setactivesheetindex(0)
                //向Excel中添加数据
                ->setCellValue('A1', 'OnSale_num_iid')
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
    
    
    //获取API属性
    public function _getAPIValue($page_no,$page_size) {
        //num_iid不存在则返回NULL
        $_itemsTmallAll = array();
        $_itemsTmall = $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'], $page_no,$page_size);
        if (!empty($_itemsTmall)) {
            if (array_key_exists('item', $_itemsTmall['items_onsale_get_response']['items'])) {
                array_push($_itemsTmallAll, $_itemsTmall['items_onsale_get_response']['items']['item']);
            }
            return $_itemsTmallAll;
        } else {
            return $_itemsTmall;
        }
    }
    
     public function _getTotalResults($page_no,$page_size) {
        $_itemsTmallNoAll = array();
        
        $_itemsTmallNo = $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'],$page_no,$page_size);
        if (!empty($_itemsTmallNo)) {
            if (array_key_exists('total_results', $_itemsTmallNo['items_onsale_get_response'])) {
                array_push($_itemsTmallNoAll, $_itemsTmallNo['items_onsale_get_response']['total_results']);
            }
            //total_results=37
            return $_itemsTmallNoAll[0];
        } else {
            return $_itemsTmallNo;
        }
    }
    //连接淘宝天猫API
    private function _connectTmall($sessionkey,$page_no,$page_size) {
        $_taobaoConnect = new TaobaoConnectorOnSale();
        $_taobaoConnect->__url = Yii::app()->params['taobao_api']['url'];
        $_taobaoConnect->__appkey = Yii::app()->params['taobao_api']['appkey'];
        $_taobaoConnect->__appsecret = Yii::app()->params['taobao_api']['appsecret'];
        $_taobaoConnect->__method = Yii::app()->params['taobao_api']['method4'];
        $_taobaoConnect->__fields = Yii::app()->params['taobao_api']['fields4'];
        $_items = $_taobaoConnect->connectTaobaoonsale($sessionkey, $page_no,$page_size);
        if (array_key_exists('error_response', $_items)) {
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
            return NULL;
        }
        if (array_key_exists('items_onsale_get_response', $_items)) {
            if (!empty($_items)) {
                return $_items;
            } else {
                Yii::log('No data numiid', 'error', 'system.fail');
                return NULL;
            }
        } else {
            return NULL;
        }
    }
    
}