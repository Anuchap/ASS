using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace Core
{
    public class ExcelConfiguration
    {
        public Dictionary<string, string> Discipline { get; set; }

        public Dictionary<string, string> Display { get; set; }

        public Dictionary<string, string> Vdo { get; set; }

        public Dictionary<string, string> YouTube { get; set; }

        public Dictionary<string, string> Facebook { get; set; }

        public Dictionary<string, string> Instagram { get; set; }

        public Dictionary<string, string> Creative { get; set; }

        public ExcelConfiguration()
        {
            //Discipline = new Dictionary<string, string>
            //{
            //    {"Search", "B"},
            //    {"Display", "C"},
            //    {"Online Video", "G"},
            //    {"YouTube Ad", "K"},
            //    {"Facebook Ad", "Q"},
            //    {"Instagram Ad", "W"},
            //    {"Twitter Ad", "Z"},
            //    {"Line", "AA"},
            //    {"Instant Messaging", "AB"},
            //    {"Social", "AC"},
            //    {"Creative", "AD"},
            //    {"Other", "AH"}
            //};

            //Display = new Dictionary<string, string>
            //{
            //    {"Direct", "C"},
            //    {"Ad Network", "D"},
            //    {"Programmatic", "E"}
            //};

            //OnlineVideo = new Dictionary<string, string>
            //{
            //    {"Direct", "G"},
            //    {"Ad Network", "H"},
            //    {"Programmatic", "I"}
            //};

            //YouTubeAd = new Dictionary<string, string>
            //{
            //    {"Display Desktop", "K"},
            //    {"Display Mobile", "L"},
            //    {"Video Desktop", "N"},
            //    {"Video Mobile", "O"}
            //};

            //FacebookAd = new Dictionary<string, string>
            //{
            //    {"Display Desktop", "Q"},
            //    {"Display Mobile", "R"},
            //    {"Video Desktop", "T"},
            //    {"Video Mobile", "U"}
            //};

            //InstagramAd = new Dictionary<string, string>
            //{
            //    {"Display", "W"},
            //    {"Video", "X"}
            //};

            //Creative = new Dictionary<string, string>
            //{
            //    {"Online Video Production", "AD"},
            //    {"Web Banner & App Production", "AE"},
            //    {"Social Media Platform Management", "AF"}
            //};
        }

        public static ExcelConfiguration GetConfiguration()
        {
            return JsonConvert.DeserializeObject<ExcelConfiguration>(File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "config.json"));
        }
    }
}