<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
Session_Start();
$_SESSION["finger_id"] = "";
$_SESSION["device_id"] = "";
$_SESSION["activle_user"] = "";
$_SESSION["active_user_zh"] = "";

$finger_id = $_POST['finger_id'];
$device_id = $_POST['device_id'];

if ($device_id != "abc987") {
    echo "wrong device";
    return;
}

/*
  jason <jason.hsu@skyrock-casting.com>,
  711 FC/孙永飞 <yf.sun@skyrockcasting.com>,
  712 wuwenqing <wenqing.wu@fulltech-metal.com>,
  713 wwy.wu <wwy.wu@fulltech-metal.com>,
  714 vicky <vicky.li@skyrock-casting.com>,
  715 富甲Harvey <harvey.zhu@skyrock-casting.com>,
 */


$nameList[611] = "test111";
$nameList[612] = "test222";
$nameList[613] = "test333";
$nameZhList[611] = "測試帳號一";
$nameZhList[612] = "測試帳號二";
$nameZhList[613] = "測試帳號三";

$nameList[711] = "yf.sun";
$nameList[712] = "wenqing.wu";
$nameList[713] = "wwy.wu";
$nameList[714] = "vicky.li";
$nameList[715] = "harvey.zhu";
$nameZhList[711] = "孙永飞";
$nameZhList[712] = "吳文清";
$nameZhList[713] = "吳文益";
$nameZhList[714] = "李兰英";
$nameZhList[715] = "朱中行";





$active_user = $nameList[$finger_id];
$active_user_zh = $nameZhList[$finger_id];

if (strlen($active_user) > 0) {
    $_SESSION["finger_id"] = $finger_id;
    $_SESSION["active_user"] = $active_user;
    $_SESSION["active_user_zh"] = $active_user_zh;

    echo "ok999";
} else {
    echo "wrong finger";
}

//$pass=$_POST['pass'];
//echo "id=$id, name=$name, active_user=$active_user";

//echo "ok777";
