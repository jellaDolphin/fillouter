
<?php

require("../database.php");
require("../functions.php");

header('Access-Control-Allow-Origin:*', TRUE, 200);

switch($_SERVER['REQUEST_METHOD']) {

    case 'GET': $the_request = &$_GET; break;
    case 'POST': $the_request = &$_POST; break;
    default: $the_request = "";
}

if($the_request) {
   
    global $GPath;
    //Authentification
    if(isset($the_request["stoken"]) && isset($the_request["ptoken"])){
        $token = $the_request["stoken"];
        $plugin_token = $the_request["ptoken"];
        if(user_verified($token, $plugin_token) == "success");
        else {
            echo json_encode(array("status"=>"failed"));
            return;
        }
    }
    else if(isset($the_request["stoken"]) && isset($the_request["disconnect"])){
        $token = $the_request["stoken"];
        $status = user_remove_plugin($token);
        if($status == "success")
            echo json_encode(array("status"=> "disconnect"));
        else
            echo json_encode(array("status"=>"failed"));
        return;
    }
    else if(isset($the_request["stoken"])){
        $token = $the_request["stoken"];
        $plugin_token = bin2hex(openssl_random_pseudo_bytes(16));
        $status = user_plugin($token, $plugin_token);
        if($status == "success")
            echo json_encode(array("status"=> $status, "plugin_token"=> $plugin_token));
        else
            echo json_encode(array("status"=>"failed"));
        return;
    }
    else {
        echo json_encode(array("status"=>"failed"));
        return;
    }
    //json progress
    
    if (isset($the_request["templates"])){

        echo json_encode(get_templates());
    }
    else if (isset($the_request["newfolder"])){

        $newfolder = $the_request["newfolder"];
        if($newfolder && !is_dir($GPath.$newfolder))
            mkdir($GPath.$newfolder);
        echo json_encode(array("status" => "success"));
    }
    else if (isset($the_request["template"])){

        $template_id = $the_request["template"];
        echo json_encode(get_template($template_id));
    }
    else if (isset($the_request["cars"])){

        echo json_encode(get_cars());
    }
    else if (isset($the_request["car"])){

        $car_id = $the_request["car"];
        echo json_encode(get_car($car_id));
    }
    else if (isset($the_request["template_save"])) {
        
        $json_data = json_decode($the_request["template_save"]);
        
        
        if(!$json_data)
        {
            echo json_encode(array("status"=>"error", "buf"=>"ERROR"));
            return;
        }
        
        $buf = template_save($json_data);
        // return $buf;

        if($buf != "success"){
            echo json_encode(array("status"=>"error", "buf"=>$buf));
            return;
        }
        $return_data = array("status" => "success", "buf" => $buf);

        echo json_encode($return_data);
    }

    else if(isset($the_request["delete_file"])){

        $fname = $the_request["delete_file"];
        $return_data = array("status" => delete_file($fname));

        echo json_encode($return_data);
    }
    else if(isset($the_request["copy_file"])){

        $id = $the_request["copy_file"];
        $array = copy_file($id);
        $return_data = array("status" => "success", "id" => $array[0], "name" => $array[1]);

        echo json_encode($return_data);
    }
    else if(isset($the_request["delete_dir"])){

        $dname = $the_request["delete_dir"];
        $status = delete_dir($dname);
        if($status == true)
            $status = "success";
        else
            $status = "failed";
        $return_data = array("status" => $status);

        echo json_encode($return_data);
    }
}

$GConn->close();