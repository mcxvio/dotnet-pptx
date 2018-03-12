﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using dotnetpptx.Service;

namespace dotnet_pptx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string filepath = @"PresentationFromFilename.pptx";

            PptxService.CreatePresentation(filepath);
            
            BuildWebHost(args).Run();
        }

        public static IWebHost BuildWebHost(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
                .UseStartup<Startup>()
                .Build();
    }
}