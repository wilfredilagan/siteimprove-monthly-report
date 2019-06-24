<?php

require 'vendor/autoload.php';
set_time_limit(0);

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;


$site_id_array = array("xxxxxxxxxxxxx" => "xxxxxxxxxxxxxxxx", "xxxxxxxxxx" => "xxxxxxxxxxxx", "xxxxxxxxxxxxxx" => "xxxxxxxxxxxxxx", "xxxxxxxxxxxxxx" => "xxxxxxxxxxxxx","xxxxxxxxxxxxxx" => "xxxxxxxxxxxx","xxxxxxxxxxxxxx" => "xxxxxxxxxxxx");

$sheet_month = date('m');
if ($sheet_month < 11 ){ 
    $sheet_prev_month = $sheet_month - 1;
    $sheet_prev_month = "0". $sheet_prev_month;
} elseif ($sheet_month == "01") {
    $sheet_prev_month = 12;
} else{
    $sheet_prev_month = $sheet_month - 1;
}

//Login details
$api_user = "xxxxxxxxxxxxxxxxxxx";
$api_key = "xxxxxxxxxxxxxxxxxxxxxxxx";
$month = date('n')+3;
$alphabet = range('A', 'Z');
$alpha_column = $alphabet[$month-1];

foreach ($site_id_array as $site_name => $site_id){

    print "Site: " . $site_name . "</br>";
    flush();
    ob_flush();

    //Load spreadsheet
    $reader = IOFactory::createReader('Xlsx');
    $reader->setIncludeCharts(true);
    $spreadsheet = $reader->load("./monthly/{$site_name}_{$sheet_prev_month}_2018_Total_Monthly_Report.xlsx");


    //print "<pre>";
    //print_r($month);
    //print "</pre>";

    //Get all issues for site
    $url = "https://api.siteimprove.com/v2/sites/{$site_id}/accessibility/issues?page=1&page_size=200";
    $issue_json = cURL($url, $api_user, $api_key);

    $cellValue = "test";

    for ($row = 21; $cellValue != null; $row++){//Look at existing column for matching issue title
        $cellValue = $spreadsheet->getActiveSheet()->getCellbyColumnandRow(1, $row)->getValue();
        $total_num_issues = $issue_json['total_items'];
        foreach ($issue_json['items'] as $item){
            
            if ($item['conformance_level'] ==  'aaa'){//exclude AAA issues
                continue;
            };

            $error_title = $item['help']['title'];
            $error_title = str_replace("\n","",$error_title);
            $error_title = str_replace("\r","",$error_title);
            $check_id = $item['id'];
            $criterion = $item["success_criterion"];
            
            if ($cellValue == $error_title){
                // API call to get instances of issue   
                $url = "https://api.siteimprove.com/v2/sites/{$site_id}/accessibility/issues/{$criterion}/{$check_id}/progress/history?period=last_seven_days";
            
                $stats_json = cURL($url, $api_user, $api_key);
                
                $total_num_stats = $stats_json['total_items'] - 1;

                $instances = $stats_json['items'][$total_num_stats]['instances_of_this_issue'];
                

                $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month, $row, $instances);
                $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16, $row, "={$alpha_column}{$row} - D{$row}");
                $printed = true;
                print "Issue : " . $error_title . " - Instances: " . $instances . "</br>";
                flush();
                ob_flush();
                $match[]= $error_title;// create an array with all titles that exist in spreadsheet
                continue;
            }
        } 
        
    };

    foreach ($issue_json['items'] as $org_issue){// create an array with all issue titles
        $error_title = $org_issue['help']['title'];
        $originalIssue[] = $error_title;
    }

    $match = array_unique($match);
    $originalIssue = array_unique($originalIssue);
    $final_array = array_diff($originalIssue, $match); // remove all matches from arrays


    $row = $row - 1;

    foreach ($issue_json['items'] as $issue_items){
        
        foreach ($final_array as $issues){  
            $conformance = $issue_items['conformance_level'];
            
            if ($conformance ==  'aaa'){
                continue;
            };

            $severity = $issue_items['severity'];
            $error_title = $issue_items['help']['title'];
            $error_title = str_replace("\n","",$error_title);
            $error_title = str_replace("\r","",$error_title);
            $check_id = $issue_items['id'];
            $criterion = $issue_items["success_criterion"];
            //Get instances of issues that are new this month
            if ($error_title == $issues and $issue_items['pages'] != 0){

                $url = "https://api.siteimprove.com/v2/sites/{$site_id}/accessibility/issues/{$criterion}/{$check_id}/progress/history?period=last_seven_days";
            
                $stats_json = cURL($url, $api_user, $api_key);
                
                $total_num_stats = $stats_json['total_items'] - 1;

                $instances = $stats_json['items'][$total_num_stats]['instances_of_this_issue'];
                $spreadsheet->getActiveSheet()->insertNewRowBefore($row,1);
                $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(1,$row,$issues);
                $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(2, $row, strtoupper($conformance));            
                $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(3, $row, ucwords($severity));
                for ($temp = 4; $temp < $month; $temp++){
                    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($temp, $row, '0');
                }
                $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16, $row, "={$alpha_column}{$row} - D{$row}");
                $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month, $row, $instances);
                $printed = true;
                print "Issue : " . $error_title . " - Instances: " . $instances . "</br>";
                $row++;
                flush();
                ob_flush();
                continue;
            }
        }
    }
    //Get total instances of that column
    $adj_row = $row - 1;
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,$row, "=SUM({$alpha_column}21:{$alpha_column}{$adj_row})");
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16,$row, "=SUM(P21:P{$adj_row})");

    $url = "https://api.siteimprove.com/v2/sites/{$site_id}/accessibility/overview/progress/instances/history?period=last_seven_days";
    $total_instances_json = cURL($url, $api_user, $api_key);
    $num_instances = $total_instances_json['total_items'] - 1;
    $total_pages = $total_instances_json['items'][$num_instances]['total_pages'];
    $a_instances = $total_instances_json['items'][$num_instances]['a_instances'];
    $aa_instances = $total_instances_json['items'][$num_instances]['aa_instances'];
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,10,$a_instances);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16,10, "={$alpha_column}10 - D10");
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,11,$aa_instances);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16,11, "={$alpha_column}11 - D11");
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,3,$total_pages);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16, 3, "={$alpha_column}3 - D3");


    $url = "https://api.siteimprove.com/v2/sites/{$site_id}/accessibility/overview/progress/issues/history?period=last_seven_days";
    $total_issues_json = cURL($url, $api_user, $api_key);
    $num_issues = $total_issues_json['total_items'] - 1;
    $a_issues = $total_issues_json['items'][$num_issues]['a_issues'];
    $aa_issues = $total_issues_json['items'][$num_issues]['aa_issues'];
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,6,$a_issues);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16, 6, "={$alpha_column}6 - D6");
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,7,$aa_issues);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16, 7, "={$alpha_column}7 - D7");



    $url = "https://api.siteimprove.com/v2/sites/{$site_id}/quality_assurance/overview/check_history";
    $total_qa_json = cURL($url, $api_user, $api_key);
    $broken_links = $total_qa_json['items'][0]['broken_links'];
    $misspellings = $total_qa_json['items'][0]['misspellings'];
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,13,$broken_links);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16, 13, "={$alpha_column}13 - D13");
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,14,$misspellings);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16,14, "={$alpha_column}14 - D14");



    $url = "https://api.siteimprove.com/v2/sites/{$site_id}/accessibility/validation/pdf";
    $total_pdf_json = cURL($url, $api_user, $api_key);
    $pdf_errors = $total_pdf_json['total_items'];
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($month,16,$pdf_errors);
    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(16,16, "={$alpha_column}16 - D16");



    print "done";

        
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->setIncludeCharts(true);
    $writer->save("./monthly/{$site_name}_{$sheet_month}_2018_Total_Monthly_Report.xlsx");

}
function cURL ($url, $api_user, $api_key){
    $process_issue = curl_init($url);
    curl_setopt($process_issue, CURLOPT_SSL_VERIFYPEER, false);
    curl_setopt($process_issue, CURLOPT_SSL_VERIFYHOST, false);
    curl_setopt($process_issue, CURLOPT_HTTPHEADER, array('Accept: application/json'));
    curl_setopt($process_issue, CURLOPT_USERPWD, $api_user . ":" . $api_key);
    curl_setopt($process_issue, CURLOPT_RETURNTRANSFER, true);
 
    $issue_response = curl_exec($process_issue);      
    curl_close($process_issue);
    $issue_json = json_decode($issue_response, true);
    return $issue_json;
}
?>