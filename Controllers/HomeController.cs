using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using pptx_creator.Models;
using builder.Services;

namespace pptx_creator.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Create()
        {
            string filepath = @"my_springboard.pptx";

            //PptxService.CreatePresentation();
            SpringboardService svc = new SpringboardService();

            svc.CreatePackage(filepath);

            return View();
        }

        public IActionResult InsertSlide()
        {
            PptxService.InsertSlide();

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
