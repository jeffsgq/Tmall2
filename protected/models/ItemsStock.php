<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Description of ItemsStock
 *
 * @author zge05
 */
class ItemsStock extends CActiveRecord{
    //put your code here
     public static function model($className=__CLASS__) {
        return parent::model($className);
    }
    
    public function tableName() {
        return 'items_stock_tbl';
    }
}
