<?php
class TaobaoConnector {
    public $__url= '' ;
    public $__appkey= '' ;
    public $__appsecret= '' ;
    public $__sessionkey='' ;
    public $__method='';
    public $__fields='';
    //对应CategoryTreeCommand
    public function connectTaobaoSQL($sessionkey,$parent_cid){
        //参数数组
        try{
            $paramArr = array(
                'app_key' => $this->__appkey,
                'session' => $sessionkey,
                'method' =>  $this->__method,
                'format' => 'json',
                'v' => '2.0',
                'sign_method'=> 'md5',
                'timestamp' => date('Y-m-d H:i:s'),
                'fields' => $this->__fields,
                'parent_cid' => $parent_cid//此处必须为parent_cid，与淘宝API相对应
            );
            return $this->_returnJsonDecode($paramArr);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }
    //对应CategoryTreeCommand
    public function connectTaobaoItem($sessionkey,$num_iid){
        //参数数组
        try{
            $paramArr = array(
                'app_key' => $this->__appkey,
                'session' => $sessionkey,
                'method' =>  $this->__method,
                'format' => 'json',
                'v' => '2.0',
                'sign_method'=> 'md5',
                'timestamp' => date('Y-m-d H:i:s'),
                'fields' => $this->__fields,
                'num_iid' => $num_iid//此处必须为num_iid，与淘宝API相对应
            );
            return $this->_returnJsonDecode($paramArr);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }
    public function connectTaobaoSKU($sessionkey,$num_iid){
        //参数数组
        try{
            $paramArr = array(
                'app_key' => $this->__appkey,
                'session' => $sessionkey,
                'method' =>  $this->__method,
                'format' => 'json',
                'v' => '2.0',
                'sign_method'=> 'md5',
                'timestamp' => date('Y-m-d H:i:s'),
                'fields' => $this->__fields,
                'num_iid' => $num_iid//此处与淘宝API相对应
            );
            return $this->_returnJsonDecode($paramArr);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }
     public function connectTaobaoinventory($sessionkey,$page_no,$page_size,$banner){
        //参数数组
        try{
            $paramArr = array(
                'app_key' => $this->__appkey,
                'session' => $sessionkey,
                'method' =>  $this->__method,
                'format' => 'json',
                'v' => '2.0',
                'sign_method'=> 'md5',
                'timestamp' => date('Y-m-d H:i:s'),
                'fields' => $this->__fields,
                'page_size' =>$page_size,
                'page_no' =>$page_no,
                'banner' =>$banner,
//              'num_iid' => $num_iid//此处 与淘宝API相对应
            );
            return $this->_returnJsonDecode($paramArr);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }
    public function connectTaobaoonsale($sessionkey,$page_no){
        //参数数组
        try{
            $paramArr = array(
                
                'app_key' => $this->__appkey,
                'session' => $sessionkey,
                'method' =>  $this->__method,
                'format' => 'json',
                'v' => '2.0',
                'sign_method'=> 'md5',
                'timestamp' => date('Y-m-d H:i:s'),
                'fields' => $this->__fields,
                'page_size' => '200',
                'page_no' =>$page_no,
//               
//                'num_iid' => $num_iid//此处与淘宝API相对应
            );
            return $this->_returnJsonDecode($paramArr);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }
    public function connectTaobaoTrade($sessionkey,$tid){
//        header("Content-Type:text/html;charset=UTF-8");
        //参数数组
        try{
            $paramArr = array(
                 'app_key' => $this->__appkey,
                 'session' => $sessionkey,
                 'method' =>  $this->__method,
                 'format' => 'json',
                 'v' => '2.0',
                 'sign_method'=> 'md5',
                 'timestamp' => date('Y-m-d H:i:s'),
                 'fields' => $this->__fields,
                 'tid' => $tid,    
            );
            return $this->_returnJsonDecode($paramArr);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }
    public function connectTaobaoTraderates($sessionkey,$start_date,$end_date,$page_no){
        //参数数组
        try{
            $paramArr = array(
                'app_key' => $this->__appkey,
                'session' => $sessionkey,
                'method' =>  $this->__method,
                'format' => 'json',
                'v' => '2.0',
                'sign_method'=> 'md5',
                'timestamp' => date('Y-m-d H:i:s'),
                'fields' => $this->__fields,
                'rate_type' => 'get',
                'role' => 'buyer',
                'use_has_next' =>'true',
                'page_size' => 2,
                'page_no' =>$page_no,
                'start_date' => $start_date,//此处必须与淘宝API相对应
                'end_date' => $end_date
            );
           return $this->_returnJsonDecode($paramArr);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }
    //返回json code
    public function _returnJsonDecode($paramArr){
        $sign = $this->_createSign($paramArr);
        $strParam = $this->_createStrParam($paramArr);
        $strParam .= 'sign='.$sign;
        $url = $this->__url.$strParam; //沙箱环境调用地址
        $ch = curl_init();
        curl_setopt($ch, CURLOPT_URL, $url);
        curl_setopt ($ch, CURLOPT_RETURNTRANSFER, 1);
        $result = curl_exec ($ch);
        curl_close ($ch);
        return json_decode($result,true);
    }
    //签名函数
    private function _createSign ($paramArr) {
        $sign = $this->__appsecret;
        ksort($paramArr);
        foreach ($paramArr as $key => $val) {
            if ($key != '' && $val != '') {
                $sign .= $key.$val;
            }
        }
        $sign.=$this->__appsecret;
        $sign = strtoupper(md5($sign));
        return $sign;
    }
    //组参函数
    private function _createStrParam ($paramArr) {
         $strParam = '';
         foreach ($paramArr as $key => $val) {
         if ($key != '' && $val != '') {
                 $strParam .= $key.'='.urlencode($val).'&';
             }
         }
         return $strParam;
    }
}
