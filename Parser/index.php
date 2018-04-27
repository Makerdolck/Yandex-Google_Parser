<center>
<B>
<?php
error_reporting(E_ALL);
set_time_limit(0);                  //max_execution_time

require_once('library/simple_html_dom.php');
require_once('library/PHPExcel.php');


$words=array();                     //from Semantic core.txt
$text="";                           //or any temporary text
$rezult_mas=array();                //
define('how_deep', '20');          //for search depth  


//--------------------------------Semantic core--------------------------------\\

$File = fopen('Semantic core.txt', 'r');

while (!feof($File))
{
    $text = fgets($File, 100);
    if ($text=="\r\n"||$text==false)
    {
    continue;
    }
    $text = str_replace("\r\n", "", $text);
    $words[]= $text;
}
$text="";
fclose($File);

//--------------------------------Parsing--------------------------------\\

foreach ($words as $text) {
    Searching_YA(str_replace(" ", "%20", $text), $rezult_mas);
    Searching_GO(str_replace(" ", "+", $text), $rezult_mas);
}
                                             //TIC&REGION
TicAndRegion($rezult_mas);
                                            //Sorting
for ($i=0; $i<count($rezult_mas); $i++)
{
    if (isset($rezult_mas[$i]["top_go"])){
        $ar_sort = array();
        foreach($rezult_mas[$i]["top_go"] as &$ar_item)
        {
            $ar_sort[] = $ar_item['position'];
        }
        array_multisort($ar_sort, SORT_ASC, $rezult_mas[$i]["top_go"]);
    }
    if (isset($rezult_mas[$i]["top_ya"])){
        $ar_sort = array();
        foreach($rezult_mas[$i]["top_ya"] as &$ar_item)
        {
            $ar_sort[] = $ar_item['position'];
        }
        array_multisort($ar_sort, SORT_ASC, $rezult_mas[$i]["top_ya"]);
    }
}

// echo"<pre>";
// print_r($rezult_mas);
// echo "</pre>";
                                            //Excel
Exceler($rezult_mas);



//--------------------------------main--------------------------------\\
//--------------------------------E.N.D--------------------------------\\
//--------------------------------main--------------------------------\\

//++++++
//--------------------------------YANDEX--------------------------------\\

function Searching_YA($text, &$rezult_mas)
{    
    $flag = false;
    $url_mas=array();
    $page = 0;
    $deep=0;

    $curl = curl_init();
    curl_setopt($curl, CURLOPT_FAILONERROR, 1);
    curl_setopt($curl, CURLOPT_FOLLOWLOCATION, 1); // allow redirects
    curl_setopt($curl, CURLOPT_TIMEOUT, 10); // times out after 4s
    curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1); // return into a variable
    curl_setopt($curl, CURLOPT_USERAGENT, "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36");//Opera/10.00 (Windows NT 5.1; U; ru) Presto/2.2.0
    curl_setopt($curl, CURLOPT_REFERER, 'https://partnersearch.yandex.ru/');
    curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, 0);// не проверять SSL сертификат
    curl_setopt($curl, CURLOPT_SSL_VERIFYHOST, 0);// не проверять Host SSL сертификата
    
    while(5)
    {
        //$html = file_get_html('https://yandex.ru/search/?lr=213&msid='.time().'.38101.22891.14426&text='.$text.'&p='.$page.'&redircnt=1470779799.1');//load_file
        /////$html = file_get_html('https://yandex.ru/search/?lr=213&&text='.$text.'&p='.$page);
        //$html = file_get_html('https://partnersearch.yandex.ru/search/?rdrnd=356590&text='.$text.'&p='.$page);

        curl_setopt($curl, CURLOPT_URL, 'https://partnersearch.yandex.ru/search/?text='.$text.'&p='.$page);//rdrnd=356590&
        $data = curl_exec($curl);

        $File = fopen('core.html', 'w');
        fwrite($File, $data);
        fclose($File);
        
        $html = file_get_html("core.html");
        
        $a_links = $html->find('a');
        if (count($a_links)<=8)
        {
            if (empty($url_mas))
                return;
            break;
        }

        foreach($a_links as $value)
        {
            $url= $value->href;
            preg_match("/^(http.?:\/\/)?([^\/]+)/i", $url, $adres);
            if (empty($adres)||strpos($adres[0], "yandex.ru")||strpos($adres[0], "yandex.net")||strpos($adres[0], "google.ru")||strpos($adres[0], "google.com"))
                continue;
            if (empty($url_mas))
            {
                $url_mas[0]=cleaning($adres[0]);

                $deep++;
                if ($deep>=how_deep)
                {
                    $flag=true;
                    break;
                }
                continue;
            }
            if ((in_array(cleaning($adres[0]), $url_mas)))
                continue;
            $url_mas[]=cleaning($adres[0]);
            
            $deep++;
            if ($deep>=how_deep)
            {
                $flag=true;
                break;
            }
        }
        if ($flag)
           break;
        
        $page++;
        sleep(rand(4, 28));
    }
    
                                //Request_top
    if (empty($rezult_mas))
    {
        for ($i=0; $i<count($url_mas); $i++)
        {
            $rezult_mas[$i]["url"] = $url_mas[$i];
            $rezult_mas[$i]["top_ya"][0] = array("request"=>str_replace("%20", " ", $text), "position"=>($i+1));
        }
        return;
    }

    for ($i=0; $i<count($url_mas); $i++)
    {
        $flag = true;
        for ($i2=0; $i2<count($rezult_mas); $i2++)
        {
            if ($rezult_mas[$i2]["url"] == $url_mas[$i])
            {
                $rezult_mas[$i2]["top_ya"][] = array("request"=>str_replace("%20", " ", $text), "position"=>($i+1));
                $flag = false;
                break;
            }
        }
        if ($flag){
            $rezult_mas[]["url"] = $url_mas[$i];
            $rezult_mas[count($rezult_mas)-1]["top_ya"][0] = array("request"=>str_replace("%20", " ", $text), "position"=>($i+1));
        }
    }
}

//++++++
//--------------------------------GOOGLE--------------------------------\\

function Searching_GO($text, &$rezult_mas)
{
    $flag = false;
    $url_mas=array();
    $page = 0;
    $deep=0;

    $curl = curl_init();
    curl_setopt($curl, CURLOPT_FAILONERROR, 1);
    curl_setopt($curl, CURLOPT_FOLLOWLOCATION, 1); // allow redirects
    curl_setopt($curl, CURLOPT_TIMEOUT, 10); // times out after 4s
    curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1); // return into a variable
    curl_setopt($curl, CURLOPT_USERAGENT, "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36");//Opera/10.00 (Windows NT 5.1; U; ru) Presto/2.2.0
    curl_setopt($curl, CURLOPT_REFERER, 'https://www.google.ru/');
    curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, 0);// не проверять SSL сертификат
    curl_setopt($curl, CURLOPT_SSL_VERIFYHOST, 0);// не проверять Host SSL сертификата


    while(5)
    {

    //$html = file_get_html('https://www.google.ru/search?ie=UTF-8&hl=ru&q='.$text.'&start='.($page*10).'&sa=N'.'&gws_rd=ssl');//load_file
    ////$html = file_get_html('https://www.google.ru/search?ie=UTF-8&hl=ru&q='.$text.'&gws_rd=ssl#q='.$text.'&newwindow=1&hl=ru&start='.($page*10));
    //
    curl_setopt($curl, CURLOPT_URL, 'https://www.google.ru/search?ie=UTF-8&hl=ru&q='.$text.'&start='.($page*10).'&sa=N'.'&gws_rd=ssl');
    $data = curl_exec($curl);

    $File = fopen('core.html', 'w');
    fwrite($File, $data);
    fclose($File);
    
    $html = file_get_html("core.html");
    //


    $a_links = $html->find('a');
    if (count($a_links)<=8)
            break;

    foreach($a_links as $key=>$value)
        {
            //
            if ((string)$value->onmousedown=="return google.arwt(this)")
                continue;
            //

            $url= $value->href;
            preg_match("/(http.?:\/\/)+([^\/]+)/i", $url, $adres);
            
            if (empty($adres)||strpos($adres[0], "yandex.ru")||strpos($adres[0], "yandex.net")
            ||strpos($adres[0], "google.ru")||strpos($adres[0], "google.com")||strpos($adres[0], "youtube.com")
            ||strpos($adres[0], "blogger.com"))
                continue;
            if (empty($url_mas))
            {
                if (strpos($adres[0], "webcache.googleusercontent.com"))
                    continue;
                $url_cache= $a_links[$key+1]->href;
                // preg_match("/(http.?:\/\/)+([^\/]+)/i", $url_cache, $adres_cache);
                // if ((empty($adres_cache))||(!strpos($adres_cache[0], "webcache.googleusercontent.com")))
                //     continue;
                
                $url_mas[0]=cleaning($adres[0]);
                $deep++;
                if ($deep>=how_deep)
                {
                    $flag=true;
                    break;
                }
                
                continue;
            }

            if (strpos($adres[0], "webcache.googleusercontent.com"))
                    continue;
            if ((in_array(cleaning($adres[0]), $url_mas))||strtolower($url_mas[count($url_mas)-1])==strtolower(cleaning($adres[0])))
                continue;
            // $url_cache= $a_links[$key+1]->href;
            // preg_match("/(http.?:\/\/)+([^\/]+)/i", $url_cache, $adres_cache);
            // if (empty($adres_cache)||!strpos($adres_cache[0], "webcache.googleusercontent.com"))
            //     continue;

            $url_mas[]=cleaning($adres[0]);
            $deep++;
            if ($deep>=how_deep)
            {
                $flag=true;
                break;
            }
        }
        if ($flag)
           break;

        $page++;
        sleep(rand(4, 24));
    }

                                //Request_top
    if (empty($rezult_mas))
    {
        for ($i=0; $i<count($url_mas); $i++)
        {
            $rezult_mas[$i]["url"] = $url_mas[$i];
            $rezult_mas[$i]["top_go"][0] = array("request"=>str_replace("+", " ", $text), "position"=>($i+1));
        }
        return;
    }
    for ($i=0; $i<count($url_mas); $i++)
    {
        $flag = true;
        for ($i2=0; $i2<count($rezult_mas); $i2++)
        {
            if ($rezult_mas[$i2]["url"] == $url_mas[$i])
            {
                $rezult_mas[$i2]["top_go"][] = array("request"=>str_replace("+", " ", $text), "position"=>($i+1));
                $flag = false;
                break;
            }
        }
        if ($flag){
            $rezult_mas[]["url"] = $url_mas[$i];
            $rezult_mas[count($rezult_mas)-1]["top_go"][0] = array("request"=>str_replace("+", " ", $text), "position"=>($i+1));
        }
    }
}

//++++++
//--------------------------------TIC&Region--------------------------------\\

function TicAndRegion(&$rezult_mas)
{
    $curl = curl_init();
    curl_setopt($curl, CURLOPT_FAILONERROR, 1);
    curl_setopt($curl, CURLOPT_FOLLOWLOCATION, 1); // allow redirects
    curl_setopt($curl, CURLOPT_TIMEOUT, 10); // times out after 4s
    curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1); // return into a variable
    curl_setopt($curl, CURLOPT_USERAGENT, "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36");//Opera/10.00 (Windows NT 5.1; U; ru) Presto/2.2.0
    curl_setopt($curl, CURLOPT_REFERER, 'https://yandex.ru/yaca/');
    curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, 0);// не проверять SSL сертификат
    curl_setopt($curl, CURLOPT_SSL_VERIFYHOST, 0);// не проверять Host SSL сертификата

    for ($i=0; $i<count($rezult_mas); $i++)
    {
        //$html = file_get_html('https://yandex.ru/yaca/cy/ch/?text='.$rezult_mas[$i]["url"]);
        //
        curl_setopt($curl, CURLOPT_URL, 'https://yandex.ru/yaca/cy/ch/?text='.$rezult_mas[$i]["url"]);
        $data = curl_exec($curl);

        $File = fopen('core.html', 'w');
        fwrite($File, $data);
        fclose($File);
        
        $html = file_get_html("core.html");
        //

        $a_links = $html->find('a');
        if (count($a_links)<=7)
        {
            if (!isset($rezult_mas[0]["tic"]))
                return;
            break;
        }
                                                                    //TIC
        $div_tic = $html->find('div[class="yaca-snippet__cy"]');
        if (!empty($div_tic))
        {
            $TIC = strval($div_tic[0]);
            $TIC = (int)preg_replace('|[^0-9]*|','',$TIC);
            $rezult_mas[$i]["tic"] = $TIC;
        }
        else
            $rezult_mas[$i]["tic"] = 0;
                                                                                    //Regions
        $div_regions = $html->find('div[class="search-tags search-tags_type_geo"]');
        if (empty($div_regions))
        {
            $rezult_mas[$i]["region"] = ""; //NULL
            continue;
        }

        $key = $div_regions[0]->find('text');
        $rezult_mas[$i]["region"] = $key[1]->_;
        foreach($rezult_mas[$i]["region"] as $value)
            $rezult_mas[$i]["region"] = $value;

        sleep(rand(4, 39));
    }

                        //Tic sorting
    $mas_sort = array();
    foreach($rezult_mas as &$value)
    {
        $mas_sort[] = $value['tic'];
    }
    array_multisort($mas_sort, SORT_DESC, $rezult_mas);

}

//++++++
//--------------------------------EXCEL--------------------------------\\

function Exceler(&$rezult_mas)
{
    $i_ar = 0; $i=2;
    $mas=array();

    $phpexcel = new PHPExcel();

    $phpexcel =  PHPExcel_IOFactory::load ("BASE.xlsx");
    $objE2007W = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel2007');
    $page = $phpexcel->setActiveSheetIndexByName("BASE");

    for ($i=2; ;$i++)
    {
        $flag = false;
        if ($i_ar>=count($rezult_mas))
            break;

        $cell_url =(string) $page -> getCellByColumnAndRow(0, $i);
        if($cell_url!=null)
            $mas[] = $cell_url; 
        else
        {
            if (!empty($mas))
            {
                for(;$i_ar<count($rezult_mas);$i_ar++)
                {
                    if (!in_array($rezult_mas[$i_ar]["url"], $mas))
                    {
                        $flag = true;
                        break;
                    }
                }
            }
            else
                $flag = true;

            if ($flag)
            {
                $page -> setCellValueByColumnAndRow(0, $i, $rezult_mas[$i_ar]["url"]);        //URL
                $page -> setCellValueByColumnAndRow(1, $i, $rezult_mas[$i_ar]["region"]);     //REGION
                $page -> setCellValueByColumnAndRow(2, $i, $rezult_mas[$i_ar]["tic"]);        //TIC
                $i_top_y=0;
                for ($i_top=0; ; $i_top++)                                                    //TOP_YANDEX_ALL
                {
                    if ($i_top_y==3)
                        break;
                    if (isset($rezult_mas[$i_ar]["top_ya"][$i_top]))
                    {
                        if ($i_top==0)
                        {
                            $old_req = $rezult_mas[$i_ar]["top_ya"][$i_top]["request"];
                            $str = $rezult_mas[$i_ar]["top_ya"][$i_top]["position"]."-".$rezult_mas[$i_ar]["top_ya"][$i_top]["request"];
                            $page -> setCellValueByColumnAndRow(3+$i_top_y, $i, $str);
                            $i_top_y++;
                            continue;
                        }
                        $_req = $rezult_mas[$i_ar]["top_ya"][$i_top]["request"];
                        if ($_req==$old_req)
                            continue;
                        $str = $rezult_mas[$i_ar]["top_ya"][$i_top]["position"]."-".$rezult_mas[$i_ar]["top_ya"][$i_top]["request"];
                        $page -> setCellValueByColumnAndRow(3+$i_top_y, $i, $str);
                        $i_top_y++;
                        $old_req=$_req;
                    }
                    else
                        break;
                }
                $i_top_g=0;
                for ($i_top=0; ; $i_top++)                                                    //TOP_GOOGLE_ALL
                {
                    if ($i_top_g==3)
                        break;
                    if (isset($rezult_mas[$i_ar]["top_go"][$i_top]))
                    {
                        if ($i_top==0)
                        {
                            $old_req = $rezult_mas[$i_ar]["top_go"][$i_top]["request"];
                            $str = $rezult_mas[$i_ar]["top_go"][$i_top]["position"]."-".$rezult_mas[$i_ar]["top_go"][$i_top]["request"];
                            $page -> setCellValueByColumnAndRow(6+$i_top_g, $i, $str);
                            $i_top_g++;
                            continue;
                        }
                        $_req = $rezult_mas[$i_ar]["top_go"][$i_top]["request"];
                        if ($_req==$old_req)
                            continue;
                        $str = $rezult_mas[$i_ar]["top_go"][$i_top]["position"]."-".$rezult_mas[$i_ar]["top_go"][$i_top]["request"];
                        $page -> setCellValueByColumnAndRow(6+$i_top_g, $i, $str);
                        $i_top_g++;
                        $old_req=$_req;
                    }
                    else
                        break;
                }
                $page -> setCellValueByColumnAndRow(9, $i, date("d.m.20y"));                  //DATE

                $i_ar++;
                continue;
            }
        }

        // $date = (string) $page -> getCellByColumnAndRow(9, $i);
        // if (!$date==date("d.m.20y", strtotime('-7 days')))
        //     continue;
    }
    

    $objE2007W->save("BASE.xlsx");
}


function cleaning($str)
{
    $str = str_replace("http:", "", $str);
    $str = str_replace("https:", "", $str);
    $str = str_replace("/", "", $str);
    $str = str_replace("www.", "", $str);
    return $str;
}

?>
</B>
</center>