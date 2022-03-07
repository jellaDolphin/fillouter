
<?php

require("global.php");

function readDocx($path){

    global $GPath;
    if (file_exists($GPath.$path)) {
        $data = file_get_contents($GPath.$path);
        return base64_encode($data);
    }
    return "";
}

function copy_file($id) {
    
    global $GPath;
    global $GConn;

    $query = "SELECT file_name, fields, created_at FROM templates WHERE id=$id";
    $rslt = $GConn->query($query);

    $copy = "";
    $fname = "temp1.docx";
    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {
            
            $fname = $r["file_name"];
            $split = explode("/", $fname);

            if(count($split) > 1){
                if(!is_dir($GPath.$split[0]))
                    mkdir($GPath.$split[0]);
                $copy = $split[0]."/copy-".$split[1];
            }
            else
                $copy = "copy-".$fname;
           copy($GPath.$fname, $GPath.$copy);

            $date = $r["created_at"];
            $fields = $r["fields"];
            
            $query = "INSERT INTO templates (file_name, fields, created_at) VALUES ('$copy', '$fields', '$date')";
        }
    }

    $rslt = $GConn->query($query);

    $new_id = $GConn->query("SELECT MAX(id) FROM templates")->fetch_assoc()["MAX(id)"];

    $query = "SELECT fields, cars_or_bikes, is_default FROM fields WHERE template_id=$id";
    $rslt = $GConn->query($query);
    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {
            
            $fields = $r["fields"];
            $cars_or_bikes = $r["cars_or_bikes"];
            $is_default = $r["is_default"];
            
            $query = "INSERT INTO fields (fields, cars_or_bikes, is_default, template_id) VALUES ('$fields', '$cars_or_bikes', '$is_default', $new_id)";

            $GConn->query($query);
        }
    }

    //UPDATE
    $query = "SELECT id FROM fields WHERE template_id=$new_id";
    $rslt = $GConn->query($query);
    $update_fields = "";
    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {
            
            if($update_fields)
                $update_fields .= ",".$r["id"];
            else
                $update_fields = $r["id"];
        }
    }

    $query = "UPDATE templates SET fields='$update_fields' WHERE id=$new_id";
    //echo $query;
    $GConn->query($query);

    return array($new_id, str_replace($GPath, "", $copy ));
}
function user_verified($token, $plugin_token){

    $user_name = "example@example.com";
    global $GConn;
    $sql = "SELECT system_token, plugin_token FROM users WHERE username='$user_name'";
    $sql_system_token = $GConn->query($sql)->fetch_assoc()["system_token"];
    $sql_plugin_token = $GConn->query($sql)->fetch_assoc()["plugin_token"];

    if($token == $sql_system_token && $plugin_token == $sql_plugin_token)
        return "success";
    else
        return "failed";
}

function user_remove_plugin($token){

    $user_name = "example@example.com";
    global $GConn;
    $query = "UPDATE users SET system_token='$token', plugin_token='_'";
    $query .= " WHERE username='$user_name'";
    if($GConn->query($query))
        return "success";
    else
        return "failed";
}

function user_plugin($token, $plugin_token){

    $user_name = "example@example.com";
    global $GConn;
    $query = "UPDATE users SET system_token='$token', plugin_token='$plugin_token'";
    $query .= " WHERE username='$user_name'";
    if($GConn->query($query))
        return "success";
    else
        return "failed";
}
function writeDocx($fname, $buf) {

    global $GPath;
    
    $split = explode("/", $fname);
    if(count($split) > 1){
        if(!is_dir($GPath.$split[0]))
            mkdir($GPath.$split[0]);
    }
    $myfile = fopen($GPath.$fname, "w") or die("Unable to open file!");

    $buf = explode(",",base64_decode($buf));
    
    foreach($buf as $k=> $v){
        
        $tmp = chr($v);
        fwrite($myfile,$tmp);
    }

    fclose($myfile);
    
    return readDocx($fname);
}
function get_templates(){
    
    global $GConn;
    global $GPath;

    $cardir = array();

    $ffs = scandir($GPath);

    unset($ffs[array_search('.', $ffs, true)]);
    unset($ffs[array_search('..', $ffs, true)]);

    if (count($ffs) < 1)
        return;

    foreach($ffs as $ff){
        if(is_dir($GPath.'/'.$ff)) $cardir[count($cardir)] = $ff;
    }

    $json_data = array();

    //templates
    $sql = "SELECT * FROM templates";
    $rslt = $GConn->query($sql);

    $json_data["templates"] = array();

    $sort_array = array();

    $json_data["dirs"] = $cardir;

    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {
            
            //$json_data["templates"][$r["id"]] = $r["file_name"];
            $sort_array[count($sort_array)] = array($r["file_name"], $r["id"]);
        }
        sort($sort_array);
        foreach($sort_array as $row){
            
            $json_data["templates"][count($json_data["templates"])]=array("id"=> $row[1], "file_name" => $row[0]); 
        }
    }

    //fields
    $sql = "SELECT * FROM fields WHERE is_default=1 AND template_id in (SELECT MIN(template_id) FROM `fields`)";
    $rslt = $GConn->query($sql);

    $json_data["fields"] = array();

    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {
            
            $json_data["fields"][$r["id"]] = array("id"=>$r["id"], "fields"=>$r["fields"], "cars_or_bikes"=>$r["cars_or_bikes"]);
        }
    }

    $json_data["status"] = "success";
    return $json_data;
}
function get_template($template_id){

    global $GConn;

    $json_data = array();

    $sql = "SELECT file_name, fields FROM templates WHERE id='$template_id'";
    $fname = $GConn->query($sql)->fetch_assoc()["file_name"];
    $fields = $GConn->query($sql)->fetch_assoc()["fields"];
    $field_arr = explode(",", $fields);

    $json_data["buf"] = readDocx($fname);

    $sql = "SELECT * FROM fields";
    $rslt = $GConn->query($sql);

    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {
            
            if(in_array($r["id"], $field_arr)){

                $json_data["fields"][$r["id"]] = array();
                
                foreach($r as $k => $v){
                    $json_data["fields"][$r["id"]][$k] = $v;
                }
            }
            
        }
    }
    
    return $json_data;
    
}
function get_cars(){

    global $GConn;

    $json_data = array();

    $sql = "SELECT * FROM cars";
    $rslt = $GConn->query($sql);

    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {
            
            $json_data[$r["id"]] = $r["title"];
        }
    }

    return $json_data;
}
function get_car($car_id){

    global $GConn;

    $json_data = array();

    $sql = "SELECT * FROM cars WHERE id=$car_id";
    
    $rslt = $GConn->query($sql);

    if($rslt->num_rows > 0){

        while($r = $rslt->fetch_assoc()) {

            foreach($r as $k => $v){
                $json_data[$k] = $v;
            }
        }
    }
    
    return $json_data;
}
function template_save($json_data){
    
    
    global $GConn;

    if(!isset($json_data->id) || !isset($json_data->buf) || !isset($json_data->fname) || !isset($json_data->date) || !isset($json_data->fields))
        return "error";
        
    $template_id = $json_data->id;
    $buf = $json_data->buf;

    $fname = $json_data->fname;

    if (strpos($fname, '.docx') !== false);
    else $fname.=".docx";
    
    
    $date = $json_data->date;
    $fields = "";
    $fieldsObj = $json_data->fields;
    $status = $json_data->status;
    $carsObj = $json_data->cars;

    if($status == "replace") {
        
        $GConn->query("DELETE FROM templates WHERE id = $template_id");
    }
    $g_fieldsId = $GConn->query("SELECT MAX(id) FROM fields")->fetch_assoc()["MAX(id)"];
    $g_templateId = $GConn->query("SELECT MAX(id) FROM templates")->fetch_assoc()["MAX(id)"] + 1;
    $insert_querys = array();

    
    if($template_id == "-1");
    else
        $g_templateId = $template_id;
    //templates & fields
    foreach($fieldsObj as $id => $row){

        $tb_names = "";
        $tb_vals = "";
        $set_vals = "";
        $insert_st = false;

        
        // if (isset($json_data->status) && ($json_data->status == "add") && !array_key_exists("status", $row)) {
        //    $row->status="add";
        // }
        if (!array_key_exists("status", $row) && $status != "replace") continue;
        foreach($row as $k => $v){

            if($k == "status" || $k == "id" || $k == "template_id") continue;

            $v = str_replace("'", "\'", $v);

            if($status == "replace" || $row->status == "add"){
                
                if($tb_names){
                    $tb_names .= ",$k";
                    $tb_vals .= ",'$v'";
                }
                else{
                    $tb_names .= "$k";
                    $tb_vals .= "'$v'";
                }
            }
            else if($row->status == "update"){
                
                if($set_vals){
                    $set_vals .= ",$k='$v'";
                }
                else
                    $set_vals .= "$k='$v'";
            }
        }
        
        $insert_st = false;
        if($status == "replace" || $row->status == "add"){

            $insert_st = true;
            // $g_fieldsId ++;

            $query = "INSERT INTO fields ";
            $query .= "($tb_names, template_id)";
            $query .= " VALUES ($tb_vals, $g_templateId)";

            $insert_querys[count($insert_querys)] = $query;
            // if($fields)
            //     $fields .= ",$g_fieldsId";
            // else
            //     $fields .= "$g_fieldsId";
        }
        else if($row->status == "delete"){
                
            $query = "DELETE FROM fields WHERE id=$id AND is_default=0";
        }
        else if($row->status == "update"){

            $query = "UPDATE fields SET ";
            $query .= $set_vals;
            $query .= " WHERE id=$id AND is_default=0";

            if($fields)
                $fields .= ",$id";
            else
                $fields .= "$id";
        }
        

        if($insert_st == false && !$GConn->query($query) ){
            echo "error";
            return;
        }
    }

    // insert new row in templates    
    $fname = str_replace("'", "\'", $fname);
    if($template_id == "-1" || $status == "replace") {
        $query = "INSERT INTO templates (id, file_name, created_at) VALUES ($g_templateId, '$fname', '$date')";
        $template_id = $g_templateId;
    }
    else {
        $query = "UPDATE templates SET file_name='$fname', created_at='$date' WHERE id='$template_id'";
    }
    if( !$GConn->query($query) ){
        echo "error";
        return;
        
    }
    
    // insert new row(for new field) in fields
    foreach($insert_querys as $key => $query) {
        if( !$GConn->query($query) ){
            echo "error";
            return;
        }
    }
    

    // get field_id for template_id
    $query = "SELECT id FROM fields WHERE template_id = $template_id";
    $result = $GConn->query($query);
    // $fields = "";
    foreach($result as $row) {
        $g_fieldsId = $row["id"];
        if($fields)
            $fields .= ",$g_fieldsId";
        else
            $fields .= "$g_fieldsId";
    }
    
    // update fields for templates

    $query = "UPDATE templates SET file_name='$fname', fields='$fields', created_at='$date' WHERE id='$template_id'";
    
    if( !$GConn->query($query) ){
        echo "error";
        return;
    }
    
    //cars
    foreach($carsObj as $id => $row){

        $tb_names = "";
        $tb_vals = "";
        $set_vals = "";

        foreach($row as $k => $v){

            if($k == "status" || $k == "id")
                continue;
            $v = str_replace("'", "\'", $v);
            if(array_key_exists("status", $row) && $row->status == "add"){
                
                if($tb_names){
                    $tb_names .= ",$k";
                    $tb_vals .= ",'$v'";
                }
                else{
                    $tb_names .= "$k";
                    $tb_vals .= "'$v'";
                }
            }
            if(array_key_exists("status", $row) && $row->status == "update"){
                
                if($set_vals){
                    $set_vals .= ",$k='$v'";
                }
                else
                    $set_vals .= "$k='$v'";
            }
        }
        
        $state = 0;
        if(array_key_exists("status", $row) && $row->status == "delete"){
                
            $state = 1;
            $query = "DELETE FROM cars WHERE id=$id";
        }
        else if(array_key_exists("status", $row) && $row->status == "update"){

            $state = 1;
            $query = "UPDATE cars SET ";
            $query .= $set_vals;
            $query .= " WHERE id=$id";
        }
        else if(array_key_exists("status", $row) && $row->status == "add") {

            $state = 1;
            $query = "INSERT INTO cars ";
            $query .= "($tb_names)";
            $query .= " VALUES ($tb_vals)";
            continue;
        }

        if($state == 1 && !$GConn->query($query) ){
            // echo $query;
            return $query;
        }
    }
    // echo "ok";
    writeDocx($fname, $buf);
    return "success";
}

function delete_file($fname){
    global $GConn;
    global $GPath;


    $sql = "DELETE FROM templates WHERE file_name='$fname'";

    if(! $GConn->query($sql) )
        return "faild";
    if(file_exists($GPath.$fname)) {
        unlink($GPath.$fname);
    } 
    return "success";
}

function delete_dir($dname){
    global $GPath;
    $my = $dname;
    $dname = $GPath.$dname;
    if (!file_exists($dname)) {
        return true;
    }

    if (!is_dir($dname)) {
        return unlink($dname);
    }

    foreach (scandir($dname) as $item) {
        if ($item == '.' || $item == '..') {
            continue;
        }

        $filename = $my . "/" . $item;
        delete_file($filename);
    }

    return rmdir($dname);
}

function init_db() {

    global $GConn;
    global $GDbObject;

    $json_data = array();

    $sql = "show tables";
    $result = $GConn->query($sql);

    if($result->num_rows > 0){
        
        while($row = $result->fetch_assoc()) {

            foreach($row as $key => $val){
                
                $tbname = $val;
                $json_data[$tbname] = array();

                $sql = "SELECT * FROM $tbname";
                $rslt = $GConn->query($sql);

                if($rslt->num_rows > 0){

                    while($r = $rslt->fetch_assoc()) {
                        $pkey = false;
                        foreach($r as $k => $v){

                            if($pkey == false)
                                $pkey = $v;
                            $json_data[$tbname][$pkey][$k] = $v;
                            if($tbname == "templates" && $k == "fname")
                                $json_data[$tbname][$pkey]["buf"] = readDocx($v);
                        }
                    }
                }

            }
        }
    }

    $GDbObject = $json_data;
    return $json_data;
}

function select_db($cars, $stoken, $dsbx){

    global $GConn;
    $json_data = array();

    $sql = "show tables";
    $result = $GConn->query($sql);

    if($result->num_rows > 0){
        
        while($row = $result->fetch_assoc()) {
            foreach($row as $key => $val){
                
                $tbname = $val;

                $json_data[$tbname] = array();

                $sql = "SELECT * FROM $tbname";

                if($tbname == "users" && $stoken && $dsbx){
                    $sql = "SELECT * FROM $tbname WHERE userid=$cars AND system_token='$stoken' AND plugin_token='$dsbx'";
                }
                if($tbname == "cars" && $cars > 0 ){
                    $sql = "SELECT * FROM $tbname WHERE id=$cars";
                }
                $rslt = $GConn->query($sql);

                if($rslt->num_rows > 0){

                    while($r = $rslt->fetch_assoc()) {
                        foreach($r as $k => $v){
                            
                            $json_data[$tbname][$k] = $v;
                        }
                    }
                }

            }
        }
    }
    
    return $json_data;
}