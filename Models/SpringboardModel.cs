using System;
using System.Collections.Generic;

namespace builder
{
    public class SpringboardModel
    {
        public Project _project { get; set; }

        public SpringboardModel()
        {
        }

        public SpringboardModel(Project project)
        {
            _project = project;
        }

        public class Project
        {
            public string Question;
            public string Teaser;
            public Area[] Areas;
            public Market[] Markets;
            public Source[] Sources;
            public WordCloud[] WordClouds;
            public WordList[] WordLists;
        }

        public class Area
        {
            public string Title;
            public Springboard[] Springboards;
        }

        // One page per springboard.
        public class Springboard
        {
            public string Title;
            public string Description;
            public string ImageUrl;
            public List<Theme> Themes;
        }

        // Source Url to be displayed as "Source" anchor tag.
        // Market used to define background colour of circle.
        public class Theme
        {
            public string Title;
            public string Text;
            public string SourceUrl;
            public string Market;
        }

        public class Market
        {
            public string Name;
        }

        // One page per word cloud
        public class WordCloud
        {
            public string Title;
            public string Url;
        }

        // One page per word list
        public class WordList
        {
            public string Title;
            public string[] Words;
        }

        public class Source
        {
            public string Url;
            public string Area;
            public string Market;
        }
    }
}
