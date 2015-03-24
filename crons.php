<?php

/* 
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
defined('YII_DEBUG') or define('YII_DEBUG',true);
// including Yii
$yii=dirname(__FILE__).'/../framework/yii.php';
require_once($yii);
// we'll use a separate config file
$configFile=dirname(__FILE__).'/protected/config/console.php';
// creating and running console application
Yii::createConsoleApplication($configFile)->run();
