<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use \App\Server;

class ReportController extends Controller
{
    public function index()
    {
        return \View::make("report");        
    }

    public function download()
    {
        $report = new \App\Report();
        $report->generate();
        return redirect()->to('/index.php');
    }
}
