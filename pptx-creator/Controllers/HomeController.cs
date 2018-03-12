using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using pptx_creator.Models;
using pptx_creator.Services;

namespace pptx_creator.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            string filepath = @"pptx_creator.pptx";

            PptxService.CreatePresentation(filepath);

            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
