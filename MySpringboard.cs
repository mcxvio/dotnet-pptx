using System;
using System.Collections.Generic;
using M = builder.SpringboardModel;

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
            M.Area[] areas = new M.Area[10];
            areas[0] = new M.Area() { Title = "area title 1", Springboards = CreateSpringboards() };
            areas[1] = new M.Area() { Title = "area title 2", Springboards = CreateSpringboards() };
            areas[2] = new M.Area() { Title = "area title 3", Springboards = CreateSpringboards() };
            areas[3] = new M.Area() { Title = "area title 4", Springboards = CreateSpringboards() };
            areas[4] = new M.Area() { Title = "area title 5", Springboards = CreateSpringboards() };
            areas[5] = new M.Area() { Title = "area title 6", Springboards = CreateSpringboards() };
            areas[6] = new M.Area() { Title = "area title 7", Springboards = CreateSpringboards() };
            areas[7] = new M.Area() { Title = "area title 8", Springboards = CreateSpringboards() };
            areas[8] = new M.Area() { Title = "area title 9", Springboards = CreateSpringboards() };
            areas[9] = new M.Area() { Title = "area title 10", Springboards = CreateSpringboards() };
            return areas;
        }

        private M.Springboard[] CreateSpringboards()
        {
            M.Springboard[] springboards = new M.Springboard[5];
            springboards[0] = new M.Springboard { Title = "springboard title 1", Description = "springboard description 1", ImageUrl = "springboard image url 1", Themes = CreateSpringboardThemes() };
            springboards[1] = new M.Springboard { Title = "springboard title 2", Description = "springboard description 2", ImageUrl = "springboard image url 2", Themes = CreateSpringboardThemes() };
            springboards[2] = new M.Springboard { Title = "springboard title 3", Description = "springboard description 3", ImageUrl = "springboard image url 3", Themes = CreateSpringboardThemes() };
            springboards[3] = new M.Springboard { Title = "springboard title 4", Description = "springboard description 4", ImageUrl = "springboard image url 4", Themes = CreateSpringboardThemes() };
            springboards[4] = new M.Springboard { Title = "springboard title 5", Description = "springboard description 5", ImageUrl = "springboard image url 5", Themes = CreateSpringboardThemes() };
            return springboards;
        }

        private List<M.Theme> CreateSpringboardThemes()
        {
            List<M.Theme> themes = new List<M.Theme>();
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 1", Text = "springboard theme text 1", SourceUrl = "springboard theme source url 1", Market = "springboard theme market 1"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 2", Text = "springboard theme text 2", SourceUrl = "springboard theme source url 2", Market = "springboard theme market 2"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 3", Text = "springboard theme text 3", SourceUrl = "springboard theme source url 3", Market = "springboard theme market 3"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 4", Text = "springboard theme text 4", SourceUrl = "springboard theme source url 4", Market = "springboard theme market 4"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 5", Text = "springboard theme text 5", SourceUrl = "springboard theme source url 5", Market = "springboard theme market 5"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 6", Text = "springboard theme text 6", SourceUrl = "springboard theme source url 6", Market = "springboard theme market 6"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 7", Text = "springboard theme text 7", SourceUrl = "springboard theme source url 7", Market = "springboard theme market 7"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 8", Text = "springboard theme text 8", SourceUrl = "springboard theme source url 8", Market = "springboard theme market 8"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 9", Text = "springboard theme text 9", SourceUrl = "springboard theme source url 9", Market = "springboard theme market 9"
            });
            themes.Add(new M.Theme
            {
                Title = "springboard theme title 10", Text = "springboard theme text 10", SourceUrl = "springboard theme source url 10", Market = "springboard theme market 10"
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
            M.Source[] sources = new M.Source[3];
            sources[0] = new M.Source { Url = "source url 1", Area = "source area 1", Market = "source market 1" };
            sources[1] = new M.Source { Url = "source url 2", Area = "source area 2", Market = "source market 2" };
            sources[2] = new M.Source { Url = "source url 3", Area = "source area 3", Market = "source market 3" };
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
            M.WordList[] wordlists = new M.WordList[3];
            wordlists[0] = new M.WordList { Title = "wordlist title 1", Words = CreateWords() };
            wordlists[1] = new M.WordList { Title = "wordlist title 2", Words = CreateWords() };
            wordlists[2] = new M.WordList { Title = "wordlist title 3", Words = CreateWords() };
            return wordlists;
        }

        private string[] CreateWords()
        {
            return new string[] { "word 1", "word 2", "word 3" };
        }
    }
}
