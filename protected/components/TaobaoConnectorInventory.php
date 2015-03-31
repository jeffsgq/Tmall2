<?php

class TaobaoConnectorInventory {
    public $__url= '' ;
    public $__appkey= '' ;
    public $__appsecret= '' ;
    public $__sessionkey='' ;
    public $__method='';
    public $__fields='';
    
    public function connectTaobaoinventory($sessionkey,$page_no,$page_size){
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
//              'num_iid' => $num_iid//此处 与淘宝API相对应
            );
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
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
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


