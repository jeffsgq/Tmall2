<?php
header("content-type:text/html;charset=utf-8");
include_once (dirname(__FILE__).'/../extensions/PHPMailer/class.phpmailer.php');
include_once (dirname(__FILE__).'/../extensions/PHPMailer/class.smtp.php');
include_once (dirname(__FILE__).'/../extensions/PHPMailer/class.pop3.php');
class Taobaomail {
    public function sendTaobaoMai($fileName,$to){
            $mail  = new PHPMailer(); 
            $mail->CharSet ="UTF-8";
            //设置stmp参数
            $mail->IsSMTP();
            $mail->SMTPAuth = true;
            $mail->SMTPKeepAlive = true;
            $mail->SMTPSecure = "ssl";
            $mail->Host = "smtp.gmail.com";
            $mail->Port = 465;
            //gmail的帐号和密码
            $mail->Username   = "zhiqiang.gesz@gmail.com";
            $mail->Password   = "zhiqiang1002";
            //设置发送方
            $mail->From = "zhiqiang.gesz@gmail.com";
            $mail->FromName = "Zhiqiang Ge";
            $mail->Subject = "Tmall API Results";
            $mail->Body = 'The tmall api results, please check the attachment in the email.'; 
            $mail->WordWrap = 50;
            $mail->MsgHTML($mail->Body); 
            //设置回复地址
            $mail->AddReplyTo("zhiqiang.gesz@gmail.com","Zhiqiang Ge");
            //添加附件
//            $fileName="trades.xls";
            $path=dirname(__FILE__).'/../../Excel/'.$fileName;
            $name=$fileName;
            $mail->AddAttachment($path,$name,$encoding='base64',$type='application/octet-stream');
            //接收方的邮箱和姓名
//            $to="zhiqiang.ge@decathlon.com";
            $mail->AddAddress($to,"");
            //使用HTML格式发送邮件
            $mail->IsHTML(true);
            if(!$mail->Send()) {
                echo "\nMailer Error: " . $mail->ErrorInfo;
            } else {
                echo "\nMessage has been sent to:".$to;
            }
    }
}
