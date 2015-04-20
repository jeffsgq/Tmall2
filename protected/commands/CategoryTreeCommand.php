<?php
Yii::import('application.components.TaobaoConnector');
require_once( dirname(__FILE__) . '/../components/ConsoleCommand.php' );
class CategoryTreeCommand extends ConsoleCommand {
    protected $_className= null ;
    protected $parentid_Array = null;
    public function init(){
        $this->_className= get_class() ;
        $this->beforeAction( $this->_className, '') ;
        //数组初始化
        $this->parentid_Array = array("0");
    }
    public function run($args){
        $this->_clearMysql('0_parentcid');
        $this->_Print();
    }
    //获取API属性
    public function _getAPIValue($parent_cid){
        $_itemsTmallAll= array();
        $_itemsTmall= $this->_connectTmall(Yii::app()->params['taobao_api']['accessToken'],$parent_cid."");
        if (array_key_exists('item_cats',$_itemsTmall['itemcats_get_response'])&&array_key_exists('item_cat',$_itemsTmall['itemcats_get_response']['item_cats'])){
            array_push($_itemsTmallAll, $_itemsTmall['itemcats_get_response']['item_cats']['item_cat']);
        }
        return $_itemsTmallAll;
    }
    //使用递归调用
    public function _insertDB($parent_cid,$i){
        //获取API属性
        $_itemsTmallAll = $this->_getAPIValue($parent_cid);
        $_items = $this->_formatArray($_itemsTmallAll);
        $_items2 = $this->_jsonNbFormat($_items);
        foreach ($_items2 as $_firstKey=>$_firstValue){
            if(strcasecmp($_firstValue['is_parent'],1)==0){
                $is_parent = "true"; 
            }else{
                $is_parent = "false"; 
            }
            $this->_insertMysql($_firstValue['cid'],$is_parent,$_firstValue['name'], $_firstValue['parent_cid']);
            if(strcasecmp($_firstValue['is_parent'],1)==0){
                array_push($this->parentid_Array, $_firstValue['cid']);
                $i = $i + 1;
            }
        }
        return $i;
    }
    //循环输出并保存在Excel中
    public function _Print(){
        //循环
        $i=1;
        $k=0;
        do{
            $parent_cid = $this->parentid_Array[$k];
            $i = $this->_insertDB($parent_cid, $i);
            $k = $k + 1;
        }while($k!=$i);
        echo 'END';
    }
    private function _formatArray($_items){
        $_itemsArray=array();
        foreach ($_items as $_firstKey=>$_firstValue){
            foreach ($_firstValue as $_secnodKey=> $_secondValue){
                $_itemsArray[] = $_secondValue;
            }
        }
        return $_itemsArray;
    }
    private function _jsonNbFormat($_items){
        foreach ($_items as $key=>$val){
            $_items[$key]['cid']= number_format($val['cid'],0,'','');
            $_items[$key]['parent_cid']= number_format($val['parent_cid'],0,'','');
        }
        return $_items;
    }
    private function _connectTmall($_sessionkey,$_parentid){
        $_taobaoConnect= new TaobaoConnector();
        $_taobaoConnect->__url=Yii::app()->params['taobao_api']['url'] ;
        $_taobaoConnect->__appkey= Yii::app()->params['taobao_api']['appkey'] ;
        $_taobaoConnect->__appsecret= Yii::app()->params['taobao_api']['appsecret'] ;
        $_taobaoConnect->__method= Yii::app()->params['taobao_api']['methods']['item_method'] ;
        $_taobaoConnect->__fields= Yii::app()->params['taobao_api']['fields']['item_field'] ;
        $_items= $_taobaoConnect->connectTaobaoSQL( $_sessionkey,$_parentid) ;
        //判断$_items数组中是否存在'error_response'，存在则为true，否则为false
        if (array_key_exists('error_response',$_items)){
            Yii::log('Caught exception: ' . serialize($_items), 'error', 'system.fail');
            exit(); 
        }
        if (array_key_exists('itemcats_get_response',$_items)){
            if (!empty($_items)){
                return $_items ;           
            }else{
                Yii::log('No data parent_cid'.$_parentid, 'error', 'system.fail');
                exit();
            }
        }else{
            return false;
        }
    }
    private function _insertMysql($cid,$is_parent,$name,$parent_cid){
        $sql_insert = "insert into 0_parentcid(c_id,is_parent,name,parent_cid) values (" . $cid . "," . $is_parent . ",'" . $name . "'," . $parent_cid . ")";
        $connection= Yii::app()->db;//建立数据库连接
        $command2 = $connection->createCommand($sql_insert);
        $command2->execute();
    }
    private function _clearMysql($tblName){
        $sql_clear = "TRUNCATE TABLE ".$tblName;
        $connection= Yii::app()->db;//建立数据库连接
        $command = $connection->createCommand($sql_clear);
        $command->execute();
    }
}

