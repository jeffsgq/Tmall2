<?php

class TaobaoConnector {
    public $__url= '' ;
    public $__appkey= '' ;
    public $__appsecret= '' ;
    public $__sessionkey='' ;
    public $__method='';
    public $__fields='';
    
    public function connectTaobao($sessionkey,$startTime,$endTime,$status,$hasNext,$pageNo){
//        header("Content-Type:text/html;charset=UTF-8");
        //参数数组
        try{
            $_pageSize="";
            if ($hasNext=="true"){
                $_pageSize='100';
            }
            $paramArr = array(
                 'app_key' => $this->__appkey,
                 'session' => $sessionkey,
                 'method' =>  $this->__method,
                 'format' => 'json',
                 'v' => '2.0',
                 'sign_method'=> 'md5',
                 'timestamp' => date('Y-m-d H:i:s'),
                 'fields' => $this->__fields,
                 'start_created'=> $startTime,
                 'end_created'=> $endTime,
                 'status'=> $status,
                 'use_has_next'=> $hasNext,
                 'page_size'=>$_pageSize,
                 'page_no'=>$pageNo                
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
            
//            $ctx = stream_context_create(
//                    array( 
//                            'http' => array( 
//                                'timeout' => 60
//                            )
//                    ) 
//            ); 
//            $i=0;
//            do{
//                if($i<4){
//                    $result =file_get_contents($url, false, $ctx);
//                    $i++;
//                }else{
//                    Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
//                    exit();
//                }
//            } while($result==false);
            return json_decode($result,true);
        }
        catch( Exception $e ) {
            Yii::log('Caught exception: ' . $e->getMessage(), 'error', 'system.fail');
            return false ;
        } 
    }

    //签名函数
    private function _createSign ($paramArr) {
//        global $appSecret;
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

