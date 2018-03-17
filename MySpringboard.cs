using System;
using System.Collections.Generic;
using builder.Models;
using M = builder.Models.SpringboardModel;

namespace builder
{
    public class MySpringboard
    {
        public SpringboardModel Model { get; }
        public M.Project Project { get { return Model._project; } }

        public MySpringboard()
        {
            M.Project project = new M.Project
            {
                Question = "project question",
                Teaser = "project teaser",
                Areas = CreateAreas(),
                Markets = CreateMarkets(),
                Sources = CreateSources(),
                WordClouds = CreateWordClouds(),
                WordLists = CreateWordLists()
            };
            Model = new SpringboardModel(project);
        }

        private M.Area[] CreateAreas()
        {
            M.Area[] areas = new M.Area[2];
            areas[0] = new M.Area() { Title = "area title 1", Springboards = CreateSpringboards() };
            areas[1] = new M.Area() { Title = "area title 2", Springboards = CreateSpringboards() };
            return areas;
        }

        private M.Springboard[] CreateSpringboards()
        {
            M.Springboard[] springboards = new M.Springboard[2];
            springboards[0] = new M.Springboard { Title = "springboard title 1", Description = "springboard description 1", ImageUrl = "springboard image url 1", Themes = CreateSpringboardThemes() };
            springboards[1] = new M.Springboard { Title = "springboard title 2", Description = "springboard description 2", ImageUrl = "springboard image url 2", Themes = CreateSpringboardThemes() };
            return springboards;
        }

        private List<M.Theme> CreateSpringboardThemes()
        {
            List<M.Theme> themes = new List<M.Theme>();
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 1",
                Text = "springboard theme text 1",
                SourceUrl = "springboard theme url 1",
                Market = "springboard theme market 1"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 2",
                Text = "springboard theme text 2",
                SourceUrl = "springboard theme source url 2",
                Market = "springboard theme market 2"
            });
            return themes;
        }

        private M.Market[] CreateMarkets()
        {
            M.Market[] markets = new M.Market[2];
            markets[0] = new M.Market { Name = "market name 1" };
            markets[1] = new M.Market { Name = "market name 2" };
            return markets;
        }

        private M.Source[] CreateSources()
        {
            M.Source[] sources = new M.Source[2];
            sources[0] = new M.Source { Url = "source url 1", Area = "source area 1", Market = "source market 1" };
            sources[1] = new M.Source { Url = "source url 2", Area = "source area 2", Market = "source market 2" };
            return sources;
        }

        private M.WordCloud[] CreateWordClouds()
        {
            M.WordCloud[] wordclouds = new M.WordCloud[2];
            wordclouds[0] = new M.WordCloud { Title = "wordcloud title 1", Url = "wordcloud url 1" };
            wordclouds[1] = new M.WordCloud { Title = "wordcloud title 2", Url = "wordcloud url 2" };
            return wordclouds;
        }

        private M.WordList[] CreateWordLists()
        {
            M.WordList[] wordlists = new M.WordList[2];
            wordlists[0] = new M.WordList { Title = "wordlist title 1", Words = CreateWords() };
            wordlists[1] = new M.WordList { Title = "wordlist title 1", Words = CreateWords() };
            return wordlists;
        }

        private string[] CreateWords()
        {
            return new string[] { "word 1", "word 2", "word 3" };
        }
    }
}
