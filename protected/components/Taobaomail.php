<?php
header("content-type:text/html;charset=utf-8");
include_once (dirname(__FILE__).'/../extensions/PHPMailer/class.phpmailer.php');
include_once (dirname(__FILE__).'/../extensions/PHPMailer/class.smtp.php');
include_once (dirname(__FILE__).'/../extensions/PHPMailer/class.pop3.php');
class Taobaomail {
    public function sendTaobaoMai($fileName){
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
            $mail->Username = Yii::app()->params['addresser']['Username'];
            $mail->Password = Yii::app()->params['addresser']['Password'];
            //设置发送方
            $mail->From = Yii::app()->params['addresser']['Username'];
            $mail->FromName = Yii::app()->params['addresser']['FromName'];
            //邮件信息
            $mail->Subject = Yii::app()->params['mallmessage']['Subject'];
            $mail->Body = Yii::app()->params['mallmessage']['Body'];
            $mail->WordWrap = 50;
            $mail->MsgHTML($mail->Body); 
            //设置回复地址
            $mail->AddReplyTo(Yii::app()->params['addresser']['Username'],Yii::app()->params['addresser']['FromName']);
            //附件
            $path=dirname(__FILE__).'/../../Excel/'.$fileName;
            $name=$fileName;
            $mail->AddAttachment($path,$name,$encoding='base64',$type='application/octet-stream');
            //接收方的邮箱和姓名
            $this->sendMulti($mail);
            //使用HTML格式发送邮件
            $mail->IsHTML(true);
            if(!$mail->Send()) {
                echo "\nSend mail failed: " . $mail->ErrorInfo;
            } else {
                echo "\n\tMail success";
            }
    }
    //发送多人
    public function sendMulti($mail){
        $recipients = Yii::app()->params['recipients'];
        foreach ($recipients as $recipientkey => $recipientvalue) {
            $mail->AddAddress($recipientvalue,NULL);
        }
    }
}
