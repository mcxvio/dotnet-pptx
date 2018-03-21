using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace builder.Services
{
    public class SpringboardService
    {
        private MySpringboard _springboard = new MySpringboard();

        // Creates a PresentationDocument.
        public void CreatePackage(string filePath)
        {
            using (PresentationDocument package = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(PresentationDocument document)
        {
            ThumbnailPart thumbnailPart1 = document.AddNewPart<ThumbnailPart>("image/jpeg", "rId2");
            GenerateThumbnailPart1Content(thumbnailPart1);

            PresentationPart presentationPart1 = document.AddPresentationPart();
            GeneratePresentationPart1Content(presentationPart1);

            NotesMasterPart notesMasterPart1 = presentationPart1.AddNewPart<NotesMasterPart>("rId8");
            GenerateNotesMasterPart1Content(notesMasterPart1);

            ThemePart themePart1 = notesMasterPart1.AddNewPart<ThemePart>("rId1");
            GenerateThemePart1Content(themePart1);

            SlidePart slidePart1 = presentationPart1.AddNewPart<SlidePart>("rId3");
            GenerateSlidePart1Content(slidePart1);

            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart1Content(slideLayoutPart1);

            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            GenerateSlideMasterPart1Content(slideMasterPart1);

            SlideLayoutPart slideLayoutPart2 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId8");
            GenerateSlideLayoutPart2Content(slideLayoutPart2);

            slideLayoutPart2.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart3 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId3");
            GenerateSlideLayoutPart3Content(slideLayoutPart3);

            slideLayoutPart3.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart4 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId7");
            GenerateSlideLayoutPart4Content(slideLayoutPart4);

            slideLayoutPart4.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart5 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId2");
            GenerateSlideLayoutPart5Content(slideLayoutPart5);

            slideLayoutPart5.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart6 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart6Content(slideLayoutPart6);

            slideLayoutPart6.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart7 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId6");
            GenerateSlideLayoutPart7Content(slideLayoutPart7);

            slideLayoutPart7.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart8 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId5");
            GenerateSlideLayoutPart8Content(slideLayoutPart8);

            slideLayoutPart8.AddPart(slideMasterPart1, "rId1");

            ThemePart themePart2 = slideMasterPart1.AddNewPart<ThemePart>("rId10");
            GenerateThemePart2Content(themePart2);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId4");

            SlideLayoutPart slideLayoutPart9 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId9");
            GenerateSlideLayoutPart9Content(slideLayoutPart9);

            slideLayoutPart9.AddPart(slideMasterPart1, "rId1");

            SlidePart slidePart2 = presentationPart1.AddNewPart<SlidePart>("rId7");
            GenerateSlidePart2Content(slidePart2);

            slidePart2.AddPart(slideLayoutPart1, "rId1");

            TableStylesPart tableStylesPart1 = presentationPart1.AddNewPart<TableStylesPart>("rId12");
            GenerateTableStylesPart1Content(tableStylesPart1);

            SlidePart slidePart3 = presentationPart1.AddNewPart<SlidePart>("rId2");
            GenerateSlidePart3Content(slidePart3);

            slidePart3.AddPart(slideLayoutPart3, "rId1");

            presentationPart1.AddPart(slideMasterPart1, "rId1");

            SlidePart slidePart4 = presentationPart1.AddNewPart<SlidePart>("rId6");
            GenerateSlidePart4Content(slidePart4);

            slidePart4.AddPart(slideLayoutPart1, "rId1");

            presentationPart1.AddPart(themePart2, "rId11");

            SlidePart slidePart5 = presentationPart1.AddNewPart<SlidePart>("rId5");
            GenerateSlidePart5Content(slidePart5);

            slidePart5.AddPart(slideLayoutPart1, "rId1");

            ViewPropertiesPart viewPropertiesPart1 = presentationPart1.AddNewPart<ViewPropertiesPart>("rId10");
            GenerateViewPropertiesPart1Content(viewPropertiesPart1);

            SlidePart slidePart6 = presentationPart1.AddNewPart<SlidePart>("rId4");
            GenerateSlidePart6Content(slidePart6);

            NotesSlidePart notesSlidePart1 = slidePart6.AddNewPart<NotesSlidePart>("rId2");
            GenerateNotesSlidePart1Content(notesSlidePart1);

            notesSlidePart1.AddPart(slidePart6, "rId2");

            notesSlidePart1.AddPart(notesMasterPart1, "rId1");

            slidePart6.AddPart(slideLayoutPart9, "rId1");

            PresentationPropertiesPart presentationPropertiesPart1 = presentationPart1.AddNewPart<PresentationPropertiesPart>("rId9");
            GeneratePresentationPropertiesPart1Content(presentationPropertiesPart1);

            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId4");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            SetPackageProperties(document);
        }

        // Generates content of thumbnailPart1.
        private void GenerateThumbnailPart1Content(ThumbnailPart thumbnailPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        // Generates content of presentationPart1.
        private void GeneratePresentationPart1Content(PresentationPart presentationPart1)
        {
            Presentation presentation1 = new Presentation() { SaveSubsetFonts = true, AutoCompressPictures = false };
            presentation1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentation1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentation1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList();
            SlideMasterId slideMasterId1 = new SlideMasterId() { Id = (UInt32Value)2147483672U, RelationshipId = "rId1" };

            slideMasterIdList1.Append(slideMasterId1);

            NotesMasterIdList notesMasterIdList1 = new NotesMasterIdList();
            NotesMasterId notesMasterId1 = new NotesMasterId() { Id = "rId8" };

            notesMasterIdList1.Append(notesMasterId1);

            SlideIdList slideIdList1 = new SlideIdList();
            SlideId slideId1 = new SlideId() { Id = (UInt32Value)266U, RelationshipId = "rId2" };
            SlideId slideId2 = new SlideId() { Id = (UInt32Value)267U, RelationshipId = "rId3" };
            SlideId slideId3 = new SlideId() { Id = (UInt32Value)265U, RelationshipId = "rId4" };
            SlideId slideId4 = new SlideId() { Id = (UInt32Value)268U, RelationshipId = "rId5" };
            SlideId slideId5 = new SlideId() { Id = (UInt32Value)269U, RelationshipId = "rId6" };
            SlideId slideId6 = new SlideId() { Id = (UInt32Value)270U, RelationshipId = "rId7" };

            slideIdList1.Append(slideId1);
            slideIdList1.Append(slideId2);
            slideIdList1.Append(slideId3);
            slideIdList1.Append(slideId4);
            slideIdList1.Append(slideId5);
            slideIdList1.Append(slideId6);
            SlideSize slideSize1 = new SlideSize() { Cx = 12192000, Cy = 6858000 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000L, Cy = 9144000L };

            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            A.DefaultParagraphProperties defaultParagraphProperties1 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { Language = "en-US" };

            defaultParagraphProperties1.Append(defaultRunProperties1);

            A.Level1ParagraphProperties level1ParagraphProperties1 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill1.Append(schemeColor1);
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties2.Append(solidFill1);
            defaultRunProperties2.Append(latinFont1);

            level1ParagraphProperties1.Append(defaultRunProperties2);

            A.Level2ParagraphProperties level2ParagraphProperties1 = new A.Level2ParagraphProperties() { LeftMargin = 457189, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill2.Append(schemeColor2);
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties3.Append(solidFill2);
            defaultRunProperties3.Append(latinFont2);

            level2ParagraphProperties1.Append(defaultRunProperties3);

            A.Level3ParagraphProperties level3ParagraphProperties1 = new A.Level3ParagraphProperties() { LeftMargin = 914377, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill3.Append(schemeColor3);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties4.Append(solidFill3);
            defaultRunProperties4.Append(latinFont3);

            level3ParagraphProperties1.Append(defaultRunProperties4);

            A.Level4ParagraphProperties level4ParagraphProperties1 = new A.Level4ParagraphProperties() { LeftMargin = 1371566, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill4.Append(schemeColor4);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties5.Append(solidFill4);
            defaultRunProperties5.Append(latinFont4);

            level4ParagraphProperties1.Append(defaultRunProperties5);

            A.Level5ParagraphProperties level5ParagraphProperties1 = new A.Level5ParagraphProperties() { LeftMargin = 1828754, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill5.Append(schemeColor5);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties6.Append(solidFill5);
            defaultRunProperties6.Append(latinFont5);

            level5ParagraphProperties1.Append(defaultRunProperties6);

            A.Level6ParagraphProperties level6ParagraphProperties1 = new A.Level6ParagraphProperties() { LeftMargin = 2285943, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill6.Append(schemeColor6);
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill6);
            defaultRunProperties7.Append(latinFont6);
            defaultRunProperties7.Append(eastAsianFont6);
            defaultRunProperties7.Append(complexScriptFont6);

            level6ParagraphProperties1.Append(defaultRunProperties7);

            A.Level7ParagraphProperties level7ParagraphProperties1 = new A.Level7ParagraphProperties() { LeftMargin = 2743131, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill7.Append(schemeColor7);
            A.LatinFont latinFont7 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties8.Append(solidFill7);
            defaultRunProperties8.Append(latinFont7);

            level7ParagraphProperties1.Append(defaultRunProperties8);

            A.Level8ParagraphProperties level8ParagraphProperties1 = new A.Level8ParagraphProperties() { LeftMargin = 3200320, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill8.Append(schemeColor8);
            A.LatinFont latinFont8 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill8);
            defaultRunProperties9.Append(latinFont8);
            defaultRunProperties9.Append(eastAsianFont8);
            defaultRunProperties9.Append(complexScriptFont8);

            level8ParagraphProperties1.Append(defaultRunProperties9);

            A.Level9ParagraphProperties level9ParagraphProperties1 = new A.Level9ParagraphProperties() { LeftMargin = 3657509, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill9.Append(schemeColor9);
            A.LatinFont latinFont9 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties10.Append(solidFill9);
            defaultRunProperties10.Append(latinFont9);

            level9ParagraphProperties1.Append(defaultRunProperties10);

            defaultTextStyle1.Append(defaultParagraphProperties1);
            defaultTextStyle1.Append(level1ParagraphProperties1);
            defaultTextStyle1.Append(level2ParagraphProperties1);
            defaultTextStyle1.Append(level3ParagraphProperties1);
            defaultTextStyle1.Append(level4ParagraphProperties1);
            defaultTextStyle1.Append(level5ParagraphProperties1);
            defaultTextStyle1.Append(level6ParagraphProperties1);
            defaultTextStyle1.Append(level7ParagraphProperties1);
            defaultTextStyle1.Append(level8ParagraphProperties1);
            defaultTextStyle1.Append(level9ParagraphProperties1);

            PresentationExtensionList presentationExtensionList1 = new PresentationExtensionList();

            PresentationExtension presentationExtension1 = new PresentationExtension() { Uri = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}" };

            P15.SlideGuideList slideGuideList1 = new P15.SlideGuideList();
            slideGuideList1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            P15.ExtendedGuide extendedGuide1 = new P15.ExtendedGuide() { Id = (UInt32Value)1U, Orientation = DirectionValues.Horizontal, Position = 2160 };

            P15.ColorType colorType1 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "A4A3A4" };

            colorType1.Append(rgbColorModelHex1);

            extendedGuide1.Append(colorType1);

            P15.ExtendedGuide extendedGuide2 = new P15.ExtendedGuide() { Id = (UInt32Value)2U, Position = 3840 };

            P15.ColorType colorType2 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "A4A3A4" };

            colorType2.Append(rgbColorModelHex2);

            extendedGuide2.Append(colorType2);

            slideGuideList1.Append(extendedGuide1);
            slideGuideList1.Append(extendedGuide2);

            presentationExtension1.Append(slideGuideList1);

            presentationExtensionList1.Append(presentationExtension1);

            presentation1.Append(slideMasterIdList1);
            presentation1.Append(notesMasterIdList1);
            presentation1.Append(slideIdList1);
            presentation1.Append(slideSize1);
            presentation1.Append(notesSize1);
            presentation1.Append(defaultTextStyle1);
            presentation1.Append(presentationExtensionList1);

            presentationPart1.Presentation = presentation1;
        }

        // Generates content of notesMasterPart1.
        private void GenerateNotesMasterPart1Content(NotesMasterPart notesMasterPart1)
        {
            NotesMaster notesMaster1 = new NotesMaster();
            notesMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            notesMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            notesMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData1 = new CommonSlideData();

            Background background1 = new Background();

            BackgroundStyleReference backgroundStyleReference1 = new BackgroundStyleReference() { Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference1.Append(schemeColor10);

            background1.Append(backgroundStyleReference1);

            ShapeTree shapeTree1 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties1 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties1 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(nonVisualGroupShapeDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            GroupShapeProperties groupShapeProperties1 = new GroupShapeProperties();

            A.TransformGroup transformGroup1 = new A.TransformGroup();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset1 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Header Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape1 = new PlaceholderShape() { Type = PlaceholderValues.Header, Size = PlaceholderSizeValues.Quarter };

            applicationNonVisualDrawingProperties2.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents() { Cx = 2971800L, Cy = 458788L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };

            A.ListStyle listStyle1 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties2 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left };
            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties2.Append(defaultRunProperties11);

            listStyle1.Append(level1ParagraphProperties2);

            A.Paragraph paragraph1 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            Shape shape2 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties2 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties3 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Date Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties3.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties3);

            ShapeProperties shapeProperties2 = new ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset3 = new A.Offset() { X = 3884613L, Y = 0L };
            A.Extents extents3 = new A.Extents() { Cx = 2971800L, Cy = 458788L };

            transform2D2.Append(offset3);
            transform2D2.Append(extents3);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);

            TextBody textBody2 = new TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };

            A.ListStyle listStyle2 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties3 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Right };
            A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties3.Append(defaultRunProperties12);

            listStyle2.Append(level1ParagraphProperties3);

            A.Paragraph paragraph2 = new A.Paragraph();

            A.Field field1 = new A.Field() { Id = "{DE027C70-1307-B348-A149-D7BA5F89E68E}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US" };
            runProperties1.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text1 = new A.Text();
            text1.Text = "3/15/2018";

            field1.Append(runProperties1);
            field1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(field1);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);

            Shape shape3 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties3 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties4 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Slide Image Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true };

            nonVisualShapeDrawingProperties3.Append(shapeLocks3);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties4 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape3 = new PlaceholderShape() { Type = PlaceholderValues.SlideImage, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties4.Append(placeholderShape3);

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties4);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties4);

            ShapeProperties shapeProperties3 = new ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset4 = new A.Offset() { X = 685800L, Y = 1143000L };
            A.Extents extents4 = new A.Extents() { Cx = 5486400L, Cy = 3086100L };

            transform2D3.Append(offset4);
            transform2D3.Append(extents4);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline() { Width = 12700 };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.PresetColor presetColor1 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill10.Append(presetColor1);

            outline1.Append(solidFill10);

            shapeProperties3.Append(transform2D3);
            shapeProperties3.Append(presetGeometry3);
            shapeProperties3.Append(noFill1);
            shapeProperties3.Append(outline1);

            TextBody textBody3 = new TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            shape3.Append(nonVisualShapeProperties3);
            shape3.Append(shapeProperties3);
            shape3.Append(textBody3);

            Shape shape4 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties4 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties5 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Notes Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties4 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks4 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties4.Append(shapeLocks4);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape4 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties5.Append(placeholderShape4);

            nonVisualShapeProperties4.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties4.Append(nonVisualShapeDrawingProperties4);
            nonVisualShapeProperties4.Append(applicationNonVisualDrawingProperties5);

            ShapeProperties shapeProperties4 = new ShapeProperties();

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset5 = new A.Offset() { X = 685800L, Y = 4400550L };
            A.Extents extents5 = new A.Extents() { Cx = 5486400L, Cy = 3600450L };

            transform2D4.Append(offset5);
            transform2D4.Append(extents5);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            shapeProperties4.Append(transform2D4);
            shapeProperties4.Append(presetGeometry4);

            TextBody textBody4 = new TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Level = 0 };

            A.Run run1 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US" };
            A.Text text2 = new A.Text();
            text2.Text = "Click to edit Master text styles";

            run1.Append(runProperties2);
            run1.Append(text2);

            paragraph4.Append(paragraphProperties1);
            paragraph4.Append(run1);

            A.Paragraph paragraph5 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Level = 1 };

            A.Run run2 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US" };
            A.Text text3 = new A.Text();
            text3.Text = "Second level";

            run2.Append(runProperties3);
            run2.Append(text3);

            paragraph5.Append(paragraphProperties2);
            paragraph5.Append(run2);

            A.Paragraph paragraph6 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties() { Level = 2 };

            A.Run run3 = new A.Run();
            A.RunProperties runProperties4 = new A.RunProperties() { Language = "en-US" };
            A.Text text4 = new A.Text();
            text4.Text = "Third level";

            run3.Append(runProperties4);
            run3.Append(text4);

            paragraph6.Append(paragraphProperties3);
            paragraph6.Append(run3);

            A.Paragraph paragraph7 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties() { Level = 3 };

            A.Run run4 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties() { Language = "en-US" };
            A.Text text5 = new A.Text();
            text5.Text = "Fourth level";

            run4.Append(runProperties5);
            run4.Append(text5);

            paragraph7.Append(paragraphProperties4);
            paragraph7.Append(run4);

            A.Paragraph paragraph8 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties() { Level = 4 };

            A.Run run5 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties() { Language = "en-US" };
            A.Text text6 = new A.Text();
            text6.Text = "Fifth level";

            run5.Append(runProperties6);
            run5.Append(text6);

            paragraph8.Append(paragraphProperties5);
            paragraph8.Append(run5);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);
            textBody4.Append(paragraph5);
            textBody4.Append(paragraph6);
            textBody4.Append(paragraph7);
            textBody4.Append(paragraph8);

            shape4.Append(nonVisualShapeProperties4);
            shape4.Append(shapeProperties4);
            shape4.Append(textBody4);

            Shape shape5 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties5 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties6 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties5 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks5 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties5.Append(shapeLocks5);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties6 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape5 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties6.Append(placeholderShape5);

            nonVisualShapeProperties5.Append(nonVisualDrawingProperties6);
            nonVisualShapeProperties5.Append(nonVisualShapeDrawingProperties5);
            nonVisualShapeProperties5.Append(applicationNonVisualDrawingProperties6);

            ShapeProperties shapeProperties5 = new ShapeProperties();

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset6 = new A.Offset() { X = 0L, Y = 8685213L };
            A.Extents extents6 = new A.Extents() { Cx = 2971800L, Cy = 458787L };

            transform2D5.Append(offset6);
            transform2D5.Append(extents6);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            shapeProperties5.Append(transform2D5);
            shapeProperties5.Append(presetGeometry5);

            TextBody textBody5 = new TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle5 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties4 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left };
            A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties4.Append(defaultRunProperties13);

            listStyle5.Append(level1ParagraphProperties4);

            A.Paragraph paragraph9 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph9.Append(endParagraphRunProperties4);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph9);

            shape5.Append(nonVisualShapeProperties5);
            shape5.Append(shapeProperties5);
            shape5.Append(textBody5);

            Shape shape6 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties6 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties7 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties6 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks6 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties6.Append(shapeLocks6);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties7 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape6 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)5U };

            applicationNonVisualDrawingProperties7.Append(placeholderShape6);

            nonVisualShapeProperties6.Append(nonVisualDrawingProperties7);
            nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);
            nonVisualShapeProperties6.Append(applicationNonVisualDrawingProperties7);

            ShapeProperties shapeProperties6 = new ShapeProperties();

            A.Transform2D transform2D6 = new A.Transform2D();
            A.Offset offset7 = new A.Offset() { X = 3884613L, Y = 8685213L };
            A.Extents extents7 = new A.Extents() { Cx = 2971800L, Cy = 458787L };

            transform2D6.Append(offset7);
            transform2D6.Append(extents7);

            A.PresetGeometry presetGeometry6 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

            presetGeometry6.Append(adjustValueList6);

            shapeProperties6.Append(transform2D6);
            shapeProperties6.Append(presetGeometry6);

            TextBody textBody6 = new TextBody();
            A.BodyProperties bodyProperties6 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle6 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties5 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Right };
            A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties5.Append(defaultRunProperties14);

            listStyle6.Append(level1ParagraphProperties5);

            A.Paragraph paragraph10 = new A.Paragraph();

            A.Field field2 = new A.Field() { Id = "{D7BFF79C-DF71-994D-85D5-70B4949C3937}", Type = "slidenum" };

            A.RunProperties runProperties7 = new A.RunProperties() { Language = "en-US" };
            runProperties7.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text7 = new A.Text();
            text7.Text = "‹#›";

            field2.Append(runProperties7);
            field2.Append(text7);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph10.Append(field2);
            paragraph10.Append(endParagraphRunProperties5);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph10);

            shape6.Append(nonVisualShapeProperties6);
            shape6.Append(shapeProperties6);
            shape6.Append(textBody6);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(shape1);
            shapeTree1.Append(shape2);
            shapeTree1.Append(shape3);
            shapeTree1.Append(shape4);
            shapeTree1.Append(shape5);
            shapeTree1.Append(shape6);

            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId1 = new P14.CreationId() { Val = (UInt32Value)705006766U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(background1);
            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);
            ColorMap colorMap1 = new ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            NotesStyle notesStyle1 = new NotesStyle();

            A.Level1ParagraphProperties level1ParagraphProperties6 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill11.Append(schemeColor11);
            A.LatinFont latinFont10 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties15.Append(solidFill11);
            defaultRunProperties15.Append(latinFont10);

            level1ParagraphProperties6.Append(defaultRunProperties15);

            A.Level2ParagraphProperties level2ParagraphProperties2 = new A.Level2ParagraphProperties() { LeftMargin = 457189, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill12 = new A.SolidFill();
            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill12.Append(schemeColor12);
            A.LatinFont latinFont11 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties16.Append(solidFill12);
            defaultRunProperties16.Append(latinFont11);

            level2ParagraphProperties2.Append(defaultRunProperties16);

            A.Level3ParagraphProperties level3ParagraphProperties2 = new A.Level3ParagraphProperties() { LeftMargin = 914377, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties17 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill13.Append(schemeColor13);
            A.LatinFont latinFont12 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties17.Append(solidFill13);
            defaultRunProperties17.Append(latinFont12);

            level3ParagraphProperties2.Append(defaultRunProperties17);

            A.Level4ParagraphProperties level4ParagraphProperties2 = new A.Level4ParagraphProperties() { LeftMargin = 1371566, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties18 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill14.Append(schemeColor14);
            A.LatinFont latinFont13 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties18.Append(solidFill14);
            defaultRunProperties18.Append(latinFont13);

            level4ParagraphProperties2.Append(defaultRunProperties18);

            A.Level5ParagraphProperties level5ParagraphProperties2 = new A.Level5ParagraphProperties() { LeftMargin = 1828754, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties19 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor15);
            A.LatinFont latinFont14 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties19.Append(solidFill15);
            defaultRunProperties19.Append(latinFont14);

            level5ParagraphProperties2.Append(defaultRunProperties19);

            A.Level6ParagraphProperties level6ParagraphProperties2 = new A.Level6ParagraphProperties() { LeftMargin = 2285943, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties20 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill16.Append(schemeColor16);
            A.LatinFont latinFont15 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties20.Append(solidFill16);
            defaultRunProperties20.Append(latinFont15);

            level6ParagraphProperties2.Append(defaultRunProperties20);

            A.Level7ParagraphProperties level7ParagraphProperties2 = new A.Level7ParagraphProperties() { LeftMargin = 2743131, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties21 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill17.Append(schemeColor17);
            A.LatinFont latinFont16 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties21.Append(solidFill17);
            defaultRunProperties21.Append(latinFont16);

            level7ParagraphProperties2.Append(defaultRunProperties21);

            A.Level8ParagraphProperties level8ParagraphProperties2 = new A.Level8ParagraphProperties() { LeftMargin = 3200320, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties22 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill18.Append(schemeColor18);
            A.LatinFont latinFont17 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties22.Append(solidFill18);
            defaultRunProperties22.Append(latinFont17);

            level8ParagraphProperties2.Append(defaultRunProperties22);

            A.Level9ParagraphProperties level9ParagraphProperties2 = new A.Level9ParagraphProperties() { LeftMargin = 3657509, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914377, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties23 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill19.Append(schemeColor19);
            A.LatinFont latinFont18 = new A.LatinFont() { Typeface = "+mn-lt" };

            defaultRunProperties23.Append(solidFill19);
            defaultRunProperties23.Append(latinFont18);

            level9ParagraphProperties2.Append(defaultRunProperties23);

            notesStyle1.Append(level1ParagraphProperties6);
            notesStyle1.Append(level2ParagraphProperties2);
            notesStyle1.Append(level3ParagraphProperties2);
            notesStyle1.Append(level4ParagraphProperties2);
            notesStyle1.Append(level5ParagraphProperties2);
            notesStyle1.Append(level6ParagraphProperties2);
            notesStyle1.Append(level7ParagraphProperties2);
            notesStyle1.Append(level8ParagraphProperties2);
            notesStyle1.Append(level9ParagraphProperties2);

            notesMaster1.Append(commonSlideData1);
            notesMaster1.Append(colorMap1);
            notesMaster1.Append(notesStyle1);

            notesMasterPart1.NotesMaster = notesMaster1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex3);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex4);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex5);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex6);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex7);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex8);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex9);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex10);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex11);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex12);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont19 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont19 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont19 = new A.ComplexScriptFont() { Typeface = "" };

            majorFont1.Append(latinFont19);
            majorFont1.Append(eastAsianFont19);
            majorFont1.Append(complexScriptFont19);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont20 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont20 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont20 = new A.ComplexScriptFont() { Typeface = "" };

            minorFont1.Append(latinFont20);
            minorFont1.Append(eastAsianFont20);
            minorFont1.Append(complexScriptFont20);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill20 = new A.SolidFill();
            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill20.Append(schemeColor20);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor21.Append(luminanceModulation1);
            schemeColor21.Append(saturationModulation1);
            schemeColor21.Append(tint1);

            gradientStop1.Append(schemeColor21);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor22.Append(luminanceModulation2);
            schemeColor22.Append(saturationModulation2);
            schemeColor22.Append(tint2);

            gradientStop2.Append(schemeColor22);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor23.Append(luminanceModulation3);
            schemeColor23.Append(saturationModulation3);
            schemeColor23.Append(tint3);

            gradientStop3.Append(schemeColor23);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor24.Append(saturationModulation4);
            schemeColor24.Append(luminanceModulation4);
            schemeColor24.Append(tint4);

            gradientStop4.Append(schemeColor24);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor25.Append(saturationModulation5);
            schemeColor25.Append(luminanceModulation5);
            schemeColor25.Append(shade1);

            gradientStop5.Append(schemeColor25);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor26.Append(luminanceModulation6);
            schemeColor26.Append(saturationModulation6);
            schemeColor26.Append(shade2);

            gradientStop6.Append(schemeColor26);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill20);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline2 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill21 = new A.SolidFill();
            A.SchemeColor schemeColor27 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill21.Append(schemeColor27);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill21);
            outline2.Append(presetDash1);
            outline2.Append(miter1);

            A.Outline outline3 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill22 = new A.SolidFill();
            A.SchemeColor schemeColor28 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill22.Append(schemeColor28);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill22);
            outline3.Append(presetDash2);
            outline3.Append(miter2);

            A.Outline outline4 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill23.Append(schemeColor29);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline4.Append(solidFill23);
            outline4.Append(presetDash3);
            outline4.Append(miter3);

            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);
            lineStyleList1.Append(outline4);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex13.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill24.Append(schemeColor30);

            A.SolidFill solidFill25 = new A.SolidFill();

            A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor31.Append(tint5);
            schemeColor31.Append(saturationModulation7);

            solidFill25.Append(schemeColor31);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor32 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor32.Append(tint6);
            schemeColor32.Append(saturationModulation8);
            schemeColor32.Append(shade3);
            schemeColor32.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor32);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor33 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor33.Append(tint7);
            schemeColor33.Append(saturationModulation9);
            schemeColor33.Append(shade4);
            schemeColor33.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor33);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor34.Append(shade5);
            schemeColor34.Append(saturationModulation10);

            gradientStop9.Append(schemeColor34);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill24);
            backgroundFillStyleList1.Append(solidFill25);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        #region Springboard Slide.
        // Generates content of slidePart1.
        private void GenerateSlidePart1Content(SlidePart slidePart1)
        {
            Slide slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData2 = new CommonSlideData();

            ShapeTree shapeTree2 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties2 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties8 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties2 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties8 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties2.Append(nonVisualDrawingProperties8);
            nonVisualGroupShapeProperties2.Append(nonVisualGroupShapeDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(applicationNonVisualDrawingProperties8);

            GroupShapeProperties groupShapeProperties2 = new GroupShapeProperties();

            A.TransformGroup transformGroup2 = new A.TransformGroup();
            A.Offset offset8 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents8 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset2 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup2.Append(offset8);
            transformGroup2.Append(extents8);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            Shape shape7 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties7 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties9 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties7 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks7 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties7.Append(shapeLocks7);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties9 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape7 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties9.Append(placeholderShape7);

            nonVisualShapeProperties7.Append(nonVisualDrawingProperties9);
            nonVisualShapeProperties7.Append(nonVisualShapeDrawingProperties7);
            nonVisualShapeProperties7.Append(applicationNonVisualDrawingProperties9);
            ShapeProperties shapeProperties7 = new ShapeProperties();

            TextBody textBody7 = new TextBody();
            A.BodyProperties bodyProperties7 = new A.BodyProperties();
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();

            A.Run run6 = new A.Run();

            A.RunProperties runProperties8 = new A.RunProperties() { Language = "en-GB" };

            A.SolidFill solidFill26 = new A.SolidFill();
            A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 };

            solidFill26.Append(schemeColor35);

            runProperties8.Append(solidFill26);
            A.Text text8 = new A.Text();
            text8.Text = _springboard.Project.Areas[0].Title; //"$Area.Title";

            run6.Append(runProperties8);
            run6.Append(text8);

            A.Run run7 = new A.Run();
            A.RunProperties runProperties9 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text9 = new A.Text();
            text9.Text = " Springboards";

            run7.Append(runProperties9);
            run7.Append(text9);
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph11.Append(run6);
            paragraph11.Append(run7);
            paragraph11.Append(endParagraphRunProperties6);

            textBody7.Append(bodyProperties7);
            textBody7.Append(listStyle7);
            textBody7.Append(paragraph11);

            shape7.Append(nonVisualShapeProperties7);
            shape7.Append(shapeProperties7);
            shape7.Append(textBody7);

            // Title & Description #1
            Shape shape8 = GenerateSpringboardTitle(_springboard.Project.Areas[0].Springboards[0].Title);
            //A.Offset offset9 = new A.Offset() { X = 695326L, Y = 5116152L };
            Shape shape9 = GenerateSpringboardDescription(_springboard.Project.Areas[0].Springboards[0].Description, 695326L, 5116152L);

            // Title & Description #2
            Shape shape10 = GenerateSpringboardTitleOffset(_springboard.Project.Areas[0].Springboards[1].Title, 2834401L, 2296391L);
            Shape shape28 = GenerateSpringboardDescriptionOffset(_springboard.Project.Areas[0].Springboards[1].Description, 2920602L, 5116152L);

            // Title & Description #3
            Shape shape11 = GenerateSpringboardTitleOffset(_springboard.Project.Areas[0].Springboards[2].Title, 5059678L, 2296391L);
            Shape shape12 = GenerateSpringboardDescriptionOffset(_springboard.Project.Areas[0].Springboards[2].Description, 5145880L, 5116152L);

            // Title & Description #4
            Shape shape13 = GenerateSpringboardTitleOffset(_springboard.Project.Areas[0].Springboards[3].Title, 7284955L, 2296391L);
            Shape shape14 = GenerateSpringboardDescriptionOffset(_springboard.Project.Areas[0].Springboards[3].Description, 7371159L, 5116152L);

            // Title & Description #5
            Shape shape15 = GenerateSpringboardTitleOffset(_springboard.Project.Areas[0].Springboards[4].Title, 9596436L, 2296391L);
            Shape shape16 = GenerateSpringboardDescriptionOffset(_springboard.Project.Areas[0].Springboards[4].Description, 9596436L, 5116152L);

            #region Teaser Sub-Heading Shape
            Shape shape17 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties17 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties19 = new NonVisualDrawingProperties() { Id = (UInt32Value)30U, Name = "Rounded Rectangle 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D3E3A9B6-93BD-4B4A-9568-4FEDB24C2EB9}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties19.Append(nonVisualDrawingPropertiesExtensionList1);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties19);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties19);

            ShapeProperties shapeProperties17 = new ShapeProperties();

            A.Transform2D transform2D12 = new A.Transform2D();
            A.Offset offset14 = new A.Offset() { X = 695324L, Y = 1596299L };
            A.Extents extents14 = new A.Extents() { Cx = 5130109L, Cy = 460058L };

            transform2D12.Append(offset14);
            transform2D12.Append(extents14);

            A.PresetGeometry presetGeometry7 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RoundRectangle };

            A.AdjustValueList adjustValueList7 = new A.AdjustValueList();
            A.ShapeGuide shapeGuide1 = new A.ShapeGuide() { Name = "adj", Formula = "val 50000" };

            adjustValueList7.Append(shapeGuide1);

            presetGeometry7.Append(adjustValueList7);

            A.SolidFill solidFill27 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex() { Val = "00B0F0" };

            solidFill27.Append(rgbColorModelHex14);

            A.Outline outline5 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline5.Append(noFill2);

            shapeProperties17.Append(transform2D12);
            shapeProperties17.Append(presetGeometry7);
            shapeProperties17.Append(solidFill27);
            shapeProperties17.Append(outline5);

            ShapeStyle shapeStyle1 = new ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade6 = new A.Shade() { Val = 50000 };

            schemeColor36.Append(shade6);

            lineReference1.Append(schemeColor36);

            A.FillReference fillReference1 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor37);

            A.EffectReference effectReference1 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor38);

            A.FontReference fontReference1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor39);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            TextBody textBody17 = new TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle17 = new A.ListStyle();

            A.Paragraph paragraph21 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties() { Level = 0, DefaultTabSize = 914400 };

            A.SpaceBefore spaceBefore1 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints1 = new A.SpacingPoints() { Val = 1000 };

            spaceBefore1.Append(spacingPoints1);

            paragraphProperties6.Append(spaceBefore1);

            A.Run run17 = new A.Run();

            A.RunProperties runProperties19 = new A.RunProperties() { Language = "en-US", FontSize = 2000, Dirty = false };

            A.SolidFill solidFill28 = new A.SolidFill();
            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill28.Append(schemeColor40);
            A.LatinFont latinFont21 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties19.Append(solidFill28);
            runProperties19.Append(latinFont21);
            A.Text text19 = new A.Text();
            text19.Text = "$";

            run17.Append(runProperties19);
            run17.Append(text19);

            A.Run run18 = new A.Run();

            A.RunProperties runProperties20 = new A.RunProperties() { Language = "en-US", FontSize = 2000, Dirty = false, SpellingError = true };

            A.SolidFill solidFill29 = new A.SolidFill();
            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill29.Append(schemeColor41);
            A.LatinFont latinFont22 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties20.Append(solidFill29);
            runProperties20.Append(latinFont22);
            A.Text text20 = new A.Text();
            text20.Text = _springboard.Project.Teaser; //"Project.Teaser";

            run18.Append(runProperties20);
            run18.Append(text20);

            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2000, Dirty = false };

            A.SolidFill solidFill30 = new A.SolidFill();
            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill30.Append(schemeColor42);
            A.LatinFont latinFont23 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties7.Append(solidFill30);
            endParagraphRunProperties7.Append(latinFont23);

            paragraph21.Append(paragraphProperties6);
            paragraph21.Append(run17);
            paragraph21.Append(run18);
            paragraph21.Append(endParagraphRunProperties7);

            textBody17.Append(bodyProperties17);
            textBody17.Append(listStyle17);
            textBody17.Append(paragraph21);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties17);
            shape17.Append(shapeStyle1);
            shape17.Append(textBody17);
            #endregion

            #region Picture Placeholders
            Shape shape18 = GenerateSpringboardPicturePlaceholderOffset(6U, "Picture Placeholder 5", "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E267D8F3-E48A-4700-AB21-28A292CD6593}\" />", 10U);

            Shape shape19 = GenerateSpringboardPicturePlaceholderOffset(8U, "Picture Placeholder 7", "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{73494BF0-70FF-4AB3-9134-962977A40301}\" />", 11U);

            Shape shape20 = GenerateSpringboardPicturePlaceholderOffset(32U, "Picture Placeholder 31", "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C1F7E11A-8B90-43C7-8E96-88F39215F946}\" />", 12U);

            Shape shape21 = GenerateSpringboardPicturePlaceholderOffset(34U, "Picture Placeholder 33", "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{54D10450-84D4-404C-9A7D-B5C8EEECBA1C}\" />", 13U);

            Shape shape22 = GenerateSpringboardPicturePlaceholderOffset(36U, "Picture Placeholder 35", "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C023F899-3365-4468-BD4F-10A31B72305D}\" />", 14U);
            #endregion

            // PictureUrl #1
            Shape shape23 = GenerateSpringboardPictureUrlOffset(20U, "TextBox 19", 810051L, 4171813L, _springboard.Project.Areas[0].Springboards[0].ImageUrl);
            // PictureUrl #2
            Shape shape24 = GenerateSpringboardPictureUrlOffset(21U, "TextBox 20", 2920603L, 4187073L, _springboard.Project.Areas[0].Springboards[1].ImageUrl);
            // PictureUrl #3
            Shape shape25 = GenerateSpringboardPictureUrlOffset(22U, "TextBox 21", 5260607L, 4187073L, _springboard.Project.Areas[0].Springboards[2].ImageUrl);
            // PictureUrl #4
            Shape shape26 = GenerateSpringboardPictureUrlOffset(23U, "TextBox 22", 7485885L, 4169234L, _springboard.Project.Areas[0].Springboards[3].ImageUrl);
            // PictureUrl #5
            Shape shape27 = GenerateSpringboardPictureUrlOffset(24U, "TextBox 23", 9711162L, 4187073L, _springboard.Project.Areas[0].Springboards[4].ImageUrl);

            shapeTree2.Append(nonVisualGroupShapeProperties2);
            shapeTree2.Append(groupShapeProperties2);
            shapeTree2.Append(shape7);
            shapeTree2.Append(shape8);
            shapeTree2.Append(shape9);
            shapeTree2.Append(shape10);
            shapeTree2.Append(shape11);
            shapeTree2.Append(shape12);
            shapeTree2.Append(shape13);
            shapeTree2.Append(shape14);
            shapeTree2.Append(shape15);
            shapeTree2.Append(shape16);
            shapeTree2.Append(shape17);
            shapeTree2.Append(shape18);
            shapeTree2.Append(shape19);
            shapeTree2.Append(shape20);
            shapeTree2.Append(shape21);
            shapeTree2.Append(shape22);
            shapeTree2.Append(shape23);
            shapeTree2.Append(shape24);
            shapeTree2.Append(shape25);
            shapeTree2.Append(shape26);
            shapeTree2.Append(shape27);
            shapeTree2.Append(shape28);

            CommonSlideDataExtensionList commonSlideDataExtensionList2 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension2 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId2 = new P14.CreationId() { Val = (UInt32Value)1859341269U };
            creationId2.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension2.Append(creationId2);

            commonSlideDataExtensionList2.Append(commonSlideDataExtension2);

            commonSlideData2.Append(shapeTree2);
            commonSlideData2.Append(commonSlideDataExtensionList2);

            ColorMapOverride colorMapOverride1 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slide1.Append(commonSlideData2);
            slide1.Append(colorMapOverride1);

            slidePart1.Slide = slide1;
        }

        private Shape GenerateSpringboardPicturePlaceholderOffset(UInt32Value placeholderId, string placeholderName, string xmlUnknownElement, UInt32Value placeholderIndex)
        {
            Shape shape18 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties18 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties20 = new NonVisualDrawingProperties() { Id = placeholderId, Name = placeholderName };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement(xmlUnknownElement);

            nonVisualDrawingPropertiesExtension2.Append(openXmlUnknownElement2);

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties20.Append(nonVisualDrawingPropertiesExtensionList2);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks17 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties18.Append(shapeLocks17);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape17 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = placeholderIndex };

            applicationNonVisualDrawingProperties20.Append(placeholderShape17);

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties20);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties20);
            ShapeProperties shapeProperties18 = new ShapeProperties();

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties18);

            return shape18;
        }

        private Shape GenerateSpringboardPictureUrlOffset(UInt32Value textboxId, string textboxName, long xOffset, long yOffset, string pictureUrl)
        {
            Shape shape23 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties23 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties25 = new NonVisualDrawingProperties() { Id = textboxId, Name = textboxName };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties23 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties25 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties23.Append(nonVisualDrawingProperties25);
            nonVisualShapeProperties23.Append(nonVisualShapeDrawingProperties23);
            nonVisualShapeProperties23.Append(applicationNonVisualDrawingProperties25);

            ShapeProperties shapeProperties23 = new ShapeProperties();

            A.Transform2D transform2D13 = new A.Transform2D();
            A.Offset offset15 = new A.Offset() { X = xOffset, Y = yOffset };
            A.Extents extents15 = new A.Extents() { Cx = 1785512L, Cy = 369332L };

            transform2D13.Append(offset15);
            transform2D13.Append(extents15);

            A.PresetGeometry presetGeometry8 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList8 = new A.AdjustValueList();

            presetGeometry8.Append(adjustValueList8);
            A.NoFill noFill3 = new A.NoFill();

            shapeProperties23.Append(transform2D13);
            shapeProperties23.Append(presetGeometry8);
            shapeProperties23.Append(noFill3);

            TextBody textBody18 = new TextBody();

            A.BodyProperties bodyProperties18 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            bodyProperties18.Append(shapeAutoFit1);
            A.ListStyle listStyle18 = new A.ListStyle();

            A.Paragraph paragraph22 = new A.Paragraph();

            /*
            A.Run run19 = new A.Run();
            A.RunProperties runProperties21 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text21 = new A.Text();
            text21.Text = "#";
            run19.Append(runProperties21);
            run19.Append(text21);
            */

            A.Run run20 = new A.Run();
            A.RunProperties runProperties22 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text22 = new A.Text();
            text22.Text = pictureUrl;

            run20.Append(runProperties22);
            run20.Append(text22);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            //paragraph22.Append(run19);
            paragraph22.Append(run20);
            paragraph22.Append(endParagraphRunProperties8);

            textBody18.Append(bodyProperties18);
            textBody18.Append(listStyle18);
            textBody18.Append(paragraph22);

            shape23.Append(nonVisualShapeProperties23);
            shape23.Append(shapeProperties23);
            shape23.Append(textBody18);

            return shape23;
        }

        private Shape GenerateSpringboardDescriptionOffset(string springboardDescription, long xOffset, long yOffset)
        {
            Shape shape12 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties12 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties14 = new NonVisualDrawingProperties() { Id = (UInt32Value)15U, Name = "Text Placeholder 14" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties12 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks12 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties12.Append(shapeLocks12);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties14 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape12 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)21U };

            applicationNonVisualDrawingProperties14.Append(placeholderShape12);

            nonVisualShapeProperties12.Append(nonVisualDrawingProperties14);
            nonVisualShapeProperties12.Append(nonVisualShapeDrawingProperties12);
            nonVisualShapeProperties12.Append(applicationNonVisualDrawingProperties14);

            ShapeProperties shapeProperties12 = new ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset11 = new A.Offset() { X = xOffset, Y = yOffset };
            A.Extents extents11 = new A.Extents() { Cx = 1986439L, Cy = 1013185L };

            transform2D9.Append(offset11);
            transform2D9.Append(extents11);

            shapeProperties12.Append(transform2D9);

            TextBody textBody12 = new TextBody();
            A.BodyProperties bodyProperties12 = new A.BodyProperties();
            A.ListStyle listStyle12 = new A.ListStyle();

            A.Paragraph paragraph16 = new A.Paragraph();

            A.Run run12 = new A.Run();
            A.RunProperties runProperties14 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text14 = new A.Text();
            text14.Text = springboardDescription;//"#2Springboard.Description";

            run12.Append(runProperties14);
            run12.Append(text14);

            paragraph16.Append(run12);

            textBody12.Append(bodyProperties12);
            textBody12.Append(listStyle12);
            textBody12.Append(paragraph16);

            shape12.Append(nonVisualShapeProperties12);
            shape12.Append(shapeProperties12);
            shape12.Append(textBody12);

            return shape12;
        }

        private Shape GenerateSpringboardDescription(string springboardDescription, long xOffset, long yOffset)
        {
            Shape shape9 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties9 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties11 = new NonVisualDrawingProperties() { Id = (UInt32Value)11U, Name = "Text Placeholder 10" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties9 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks9 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties9.Append(shapeLocks9);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties11 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape9 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)17U };

            applicationNonVisualDrawingProperties11.Append(placeholderShape9);

            nonVisualShapeProperties9.Append(nonVisualDrawingProperties11);
            nonVisualShapeProperties9.Append(nonVisualShapeDrawingProperties9);
            nonVisualShapeProperties9.Append(applicationNonVisualDrawingProperties11);

            ShapeProperties shapeProperties9 = new ShapeProperties();

            A.Transform2D transform2D7 = new A.Transform2D();
            //A.Offset offset9 = new A.Offset() { X = 695326L, Y = 5116152L };
            A.Offset offset9 = new A.Offset() { X = xOffset, Y = yOffset };
            A.Extents extents9 = new A.Extents() { Cx = 2035348L, Cy = 1013185L };

            transform2D7.Append(offset9);
            transform2D7.Append(extents9);

            shapeProperties9.Append(transform2D7);

            TextBody textBody9 = new TextBody();
            A.BodyProperties bodyProperties9 = new A.BodyProperties();
            A.ListStyle listStyle9 = new A.ListStyle();

            A.Paragraph paragraph13 = new A.Paragraph();

            A.Run run9 = new A.Run();
            A.RunProperties runProperties11 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text11 = new A.Text();
            text11.Text = springboardDescription; //"#0Springboard.Description";

            run9.Append(runProperties11);
            run9.Append(text11);

            paragraph13.Append(run9);

            textBody9.Append(bodyProperties9);
            textBody9.Append(listStyle9);
            textBody9.Append(paragraph13);

            shape9.Append(nonVisualShapeProperties9);
            shape9.Append(shapeProperties9);
            shape9.Append(textBody9);
            return shape9;
        }

        private Shape GenerateSpringboardTitleOffset(string springboardTitle, long xOffset, long yOffset)
        {
            Shape shape11 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties11 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties13 = new NonVisualDrawingProperties() { Id = (UInt32Value)14U, Name = "Text Placeholder 13" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties11 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks11 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties11.Append(shapeLocks11);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties13 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape11 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)20U };

            applicationNonVisualDrawingProperties13.Append(placeholderShape11);

            nonVisualShapeProperties11.Append(nonVisualDrawingProperties13);
            nonVisualShapeProperties11.Append(nonVisualShapeDrawingProperties11);
            nonVisualShapeProperties11.Append(applicationNonVisualDrawingProperties13);

            ShapeProperties shapeProperties11 = new ShapeProperties();

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset10 = new A.Offset() { X = xOffset, Y = yOffset };
            A.Extents extents10 = new A.Extents() { Cx = 2072642L, Cy = 628578L };

            transform2D8.Append(offset10);
            transform2D8.Append(extents10);

            shapeProperties11.Append(transform2D8);

            TextBody textBody11 = new TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties();
            A.ListStyle listStyle11 = new A.ListStyle();

            A.Paragraph paragraph15 = new A.Paragraph();

            A.Run run11 = new A.Run();
            A.RunProperties runProperties13 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text13 = new A.Text();
            text13.Text = springboardTitle; //"#2Springboard.Title";

            run11.Append(runProperties13);
            run11.Append(text13);

            paragraph15.Append(run11);

            textBody11.Append(bodyProperties11);
            textBody11.Append(listStyle11);
            textBody11.Append(paragraph15);

            shape11.Append(nonVisualShapeProperties11);
            shape11.Append(shapeProperties11);
            shape11.Append(textBody11);

            return shape11;
        }

        private Shape GenerateSpringboardTitle(string springboardTitle)
        {
            Shape shape8 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties8 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties10 = new NonVisualDrawingProperties() { Id = (UInt32Value)10U, Name = "Text Placeholder 9" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties8 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks8 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties8.Append(shapeLocks8);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties10 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape8 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)16U };

            applicationNonVisualDrawingProperties10.Append(placeholderShape8);

            nonVisualShapeProperties8.Append(nonVisualDrawingProperties10);
            nonVisualShapeProperties8.Append(nonVisualShapeDrawingProperties8);
            nonVisualShapeProperties8.Append(applicationNonVisualDrawingProperties10);
            ShapeProperties shapeProperties8 = new ShapeProperties();

            TextBody textBody8 = new TextBody();
            A.BodyProperties bodyProperties8 = new A.BodyProperties();
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph12 = new A.Paragraph();

            A.Run run8 = new A.Run();
            A.RunProperties runProperties10 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text10 = new A.Text();
            text10.Text = springboardTitle; //"#0Springboard.Title";

            run8.Append(runProperties10);
            run8.Append(text10);

            paragraph12.Append(run8);

            textBody8.Append(bodyProperties8);
            textBody8.Append(listStyle8);
            textBody8.Append(paragraph12);

            shape8.Append(nonVisualShapeProperties8);
            shape8.Append(shapeProperties8);
            shape8.Append(textBody8);
            return shape8;
        }
        #endregion

        // Generates content of slideLayoutPart1.
        private void GenerateSlideLayoutPart1Content(SlideLayoutPart slideLayoutPart1)
        {
            SlideLayout slideLayout1 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData3 = new CommonSlideData() { Name = "Springboard Summary" };

            ShapeTree shapeTree3 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties3 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties31 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties3 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties31 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties3.Append(nonVisualDrawingProperties31);
            nonVisualGroupShapeProperties3.Append(nonVisualGroupShapeDrawingProperties3);
            nonVisualGroupShapeProperties3.Append(applicationNonVisualDrawingProperties31);

            GroupShapeProperties groupShapeProperties3 = new GroupShapeProperties();

            A.TransformGroup transformGroup3 = new A.TransformGroup();
            A.Offset offset21 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents21 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset3 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents3 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup3.Append(offset21);
            transformGroup3.Append(extents21);
            transformGroup3.Append(childOffset3);
            transformGroup3.Append(childExtents3);

            groupShapeProperties3.Append(transformGroup3);

            Shape shape29 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties29 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties32 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties29 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks23 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties29.Append(shapeLocks23);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties32 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape23 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties32.Append(placeholderShape23);

            nonVisualShapeProperties29.Append(nonVisualDrawingProperties32);
            nonVisualShapeProperties29.Append(nonVisualShapeDrawingProperties29);
            nonVisualShapeProperties29.Append(applicationNonVisualDrawingProperties32);
            ShapeProperties shapeProperties29 = new ShapeProperties();

            TextBody textBody24 = new TextBody();
            A.BodyProperties bodyProperties24 = new A.BodyProperties();
            A.ListStyle listStyle24 = new A.ListStyle();

            A.Paragraph paragraph28 = new A.Paragraph();

            A.Run run30 = new A.Run();
            A.RunProperties runProperties32 = new A.RunProperties() { Language = "en-US" };
            A.Text text32 = new A.Text();
            text32.Text = "Click to edit Master title style";

            run30.Append(runProperties32);
            run30.Append(text32);

            paragraph28.Append(run30);

            textBody24.Append(bodyProperties24);
            textBody24.Append(listStyle24);
            textBody24.Append(paragraph28);

            shape29.Append(nonVisualShapeProperties29);
            shape29.Append(shapeProperties29);
            shape29.Append(textBody24);

            Shape shape30 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties30 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties33 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Picture Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties30 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks24 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties30.Append(shapeLocks24);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties33 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape24 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties33.Append(placeholderShape24);

            nonVisualShapeProperties30.Append(nonVisualDrawingProperties33);
            nonVisualShapeProperties30.Append(nonVisualShapeDrawingProperties30);
            nonVisualShapeProperties30.Append(applicationNonVisualDrawingProperties33);

            ShapeProperties shapeProperties30 = new ShapeProperties();

            A.Transform2D transform2D19 = new A.Transform2D();
            A.Offset offset22 = new A.Offset() { X = 695325L, Y = 3070442L };
            A.Extents extents22 = new A.Extents() { Cx = 1900238L, Cy = 1900238L };

            transform2D19.Append(offset22);
            transform2D19.Append(extents22);

            A.PresetGeometry presetGeometry13 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList13 = new A.AdjustValueList();

            presetGeometry13.Append(adjustValueList13);

            A.SolidFill solidFill31 = new A.SolidFill();

            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 95000 };

            schemeColor43.Append(luminanceModulation9);

            solidFill31.Append(schemeColor43);

            shapeProperties30.Append(transform2D19);
            shapeProperties30.Append(presetGeometry13);
            shapeProperties30.Append(solidFill31);

            TextBody textBody25 = new TextBody();
            A.BodyProperties bodyProperties25 = new A.BodyProperties();

            A.ListStyle listStyle25 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties7 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet1 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties24 = new A.DefaultRunProperties();

            level1ParagraphProperties7.Append(noBullet1);
            level1ParagraphProperties7.Append(defaultRunProperties24);

            listStyle25.Append(level1ParagraphProperties7);

            A.Paragraph paragraph29 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph29.Append(endParagraphRunProperties13);

            textBody25.Append(bodyProperties25);
            textBody25.Append(listStyle25);
            textBody25.Append(paragraph29);

            shape30.Append(nonVisualShapeProperties30);
            shape30.Append(shapeProperties30);
            shape30.Append(textBody25);

            Shape shape31 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties31 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties34 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Picture Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties31 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks25 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties31.Append(shapeLocks25);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties34 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape25 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties34.Append(placeholderShape25);

            nonVisualShapeProperties31.Append(nonVisualDrawingProperties34);
            nonVisualShapeProperties31.Append(nonVisualShapeDrawingProperties31);
            nonVisualShapeProperties31.Append(applicationNonVisualDrawingProperties34);

            ShapeProperties shapeProperties31 = new ShapeProperties();

            A.Transform2D transform2D20 = new A.Transform2D();
            A.Offset offset23 = new A.Offset() { X = 2920603L, Y = 3070442L };
            A.Extents extents23 = new A.Extents() { Cx = 1900238L, Cy = 1900238L };

            transform2D20.Append(offset23);
            transform2D20.Append(extents23);

            A.PresetGeometry presetGeometry14 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList14 = new A.AdjustValueList();

            presetGeometry14.Append(adjustValueList14);

            A.SolidFill solidFill32 = new A.SolidFill();

            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 95000 };

            schemeColor44.Append(luminanceModulation10);

            solidFill32.Append(schemeColor44);

            shapeProperties31.Append(transform2D20);
            shapeProperties31.Append(presetGeometry14);
            shapeProperties31.Append(solidFill32);

            TextBody textBody26 = new TextBody();
            A.BodyProperties bodyProperties26 = new A.BodyProperties();

            A.ListStyle listStyle26 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties8 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet2 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties25 = new A.DefaultRunProperties();

            level1ParagraphProperties8.Append(noBullet2);
            level1ParagraphProperties8.Append(defaultRunProperties25);

            listStyle26.Append(level1ParagraphProperties8);

            A.Paragraph paragraph30 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph30.Append(endParagraphRunProperties14);

            textBody26.Append(bodyProperties26);
            textBody26.Append(listStyle26);
            textBody26.Append(paragraph30);

            shape31.Append(nonVisualShapeProperties31);
            shape31.Append(shapeProperties31);
            shape31.Append(textBody26);

            Shape shape32 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties32 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties35 = new NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Picture Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties32 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks26 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties32.Append(shapeLocks26);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties35 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape26 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties35.Append(placeholderShape26);

            nonVisualShapeProperties32.Append(nonVisualDrawingProperties35);
            nonVisualShapeProperties32.Append(nonVisualShapeDrawingProperties32);
            nonVisualShapeProperties32.Append(applicationNonVisualDrawingProperties35);

            ShapeProperties shapeProperties32 = new ShapeProperties();

            A.Transform2D transform2D21 = new A.Transform2D();
            A.Offset offset24 = new A.Offset() { X = 5145881L, Y = 3070442L };
            A.Extents extents24 = new A.Extents() { Cx = 1900238L, Cy = 1900238L };

            transform2D21.Append(offset24);
            transform2D21.Append(extents24);

            A.PresetGeometry presetGeometry15 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList15 = new A.AdjustValueList();

            presetGeometry15.Append(adjustValueList15);

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 95000 };

            schemeColor45.Append(luminanceModulation11);

            solidFill33.Append(schemeColor45);

            shapeProperties32.Append(transform2D21);
            shapeProperties32.Append(presetGeometry15);
            shapeProperties32.Append(solidFill33);

            TextBody textBody27 = new TextBody();
            A.BodyProperties bodyProperties27 = new A.BodyProperties();

            A.ListStyle listStyle27 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties9 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet3 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties26 = new A.DefaultRunProperties();

            level1ParagraphProperties9.Append(noBullet3);
            level1ParagraphProperties9.Append(defaultRunProperties26);

            listStyle27.Append(level1ParagraphProperties9);

            A.Paragraph paragraph31 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph31.Append(endParagraphRunProperties15);

            textBody27.Append(bodyProperties27);
            textBody27.Append(listStyle27);
            textBody27.Append(paragraph31);

            shape32.Append(nonVisualShapeProperties32);
            shape32.Append(shapeProperties32);
            shape32.Append(textBody27);

            Shape shape33 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties33 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties36 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Picture Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties33 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks27 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties33.Append(shapeLocks27);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties36 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape27 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)13U };

            applicationNonVisualDrawingProperties36.Append(placeholderShape27);

            nonVisualShapeProperties33.Append(nonVisualDrawingProperties36);
            nonVisualShapeProperties33.Append(nonVisualShapeDrawingProperties33);
            nonVisualShapeProperties33.Append(applicationNonVisualDrawingProperties36);

            ShapeProperties shapeProperties33 = new ShapeProperties();

            A.Transform2D transform2D22 = new A.Transform2D();
            A.Offset offset25 = new A.Offset() { X = 7371159L, Y = 3070442L };
            A.Extents extents25 = new A.Extents() { Cx = 1900238L, Cy = 1900238L };

            transform2D22.Append(offset25);
            transform2D22.Append(extents25);

            A.PresetGeometry presetGeometry16 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList16 = new A.AdjustValueList();

            presetGeometry16.Append(adjustValueList16);

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 95000 };

            schemeColor46.Append(luminanceModulation12);

            solidFill34.Append(schemeColor46);

            shapeProperties33.Append(transform2D22);
            shapeProperties33.Append(presetGeometry16);
            shapeProperties33.Append(solidFill34);

            TextBody textBody28 = new TextBody();
            A.BodyProperties bodyProperties28 = new A.BodyProperties();

            A.ListStyle listStyle28 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties10 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet4 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties27 = new A.DefaultRunProperties();

            level1ParagraphProperties10.Append(noBullet4);
            level1ParagraphProperties10.Append(defaultRunProperties27);

            listStyle28.Append(level1ParagraphProperties10);

            A.Paragraph paragraph32 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties16 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph32.Append(endParagraphRunProperties16);

            textBody28.Append(bodyProperties28);
            textBody28.Append(listStyle28);
            textBody28.Append(paragraph32);

            shape33.Append(nonVisualShapeProperties33);
            shape33.Append(shapeProperties33);
            shape33.Append(textBody28);

            Shape shape34 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties34 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties37 = new NonVisualDrawingProperties() { Id = (UInt32Value)10U, Name = "Picture Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties34 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks28 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties34.Append(shapeLocks28);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties37 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape28 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)14U };

            applicationNonVisualDrawingProperties37.Append(placeholderShape28);

            nonVisualShapeProperties34.Append(nonVisualDrawingProperties37);
            nonVisualShapeProperties34.Append(nonVisualShapeDrawingProperties34);
            nonVisualShapeProperties34.Append(applicationNonVisualDrawingProperties37);

            ShapeProperties shapeProperties34 = new ShapeProperties();

            A.Transform2D transform2D23 = new A.Transform2D();
            A.Offset offset26 = new A.Offset() { X = 9596437L, Y = 3070442L };
            A.Extents extents26 = new A.Extents() { Cx = 1900238L, Cy = 1900238L };

            transform2D23.Append(offset26);
            transform2D23.Append(extents26);

            A.PresetGeometry presetGeometry17 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList17 = new A.AdjustValueList();

            presetGeometry17.Append(adjustValueList17);

            A.SolidFill solidFill35 = new A.SolidFill();

            A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 95000 };

            schemeColor47.Append(luminanceModulation13);

            solidFill35.Append(schemeColor47);

            shapeProperties34.Append(transform2D23);
            shapeProperties34.Append(presetGeometry17);
            shapeProperties34.Append(solidFill35);

            TextBody textBody29 = new TextBody();
            A.BodyProperties bodyProperties29 = new A.BodyProperties();

            A.ListStyle listStyle29 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties11 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet5 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties28 = new A.DefaultRunProperties();

            level1ParagraphProperties11.Append(noBullet5);
            level1ParagraphProperties11.Append(defaultRunProperties28);

            listStyle29.Append(level1ParagraphProperties11);

            A.Paragraph paragraph33 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties17 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph33.Append(endParagraphRunProperties17);

            textBody29.Append(bodyProperties29);
            textBody29.Append(listStyle29);
            textBody29.Append(paragraph33);

            shape34.Append(nonVisualShapeProperties34);
            shape34.Append(shapeProperties34);
            shape34.Append(textBody29);

            Shape shape35 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties35 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties38 = new NonVisualDrawingProperties() { Id = (UInt32Value)11U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties35 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks29 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties35.Append(shapeLocks29);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties38 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape29 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)15U };

            applicationNonVisualDrawingProperties38.Append(placeholderShape29);

            nonVisualShapeProperties35.Append(nonVisualDrawingProperties38);
            nonVisualShapeProperties35.Append(nonVisualShapeDrawingProperties35);
            nonVisualShapeProperties35.Append(applicationNonVisualDrawingProperties38);

            ShapeProperties shapeProperties35 = new ShapeProperties();

            A.Transform2D transform2D24 = new A.Transform2D();
            A.Offset offset27 = new A.Offset() { X = 695325L, Y = 1680857L };
            A.Extents extents27 = new A.Extents() { Cx = 10801350L, Cy = 470061L };

            transform2D24.Append(offset27);
            transform2D24.Append(extents27);

            shapeProperties35.Append(transform2D24);

            TextBody textBody30 = new TextBody();
            A.BodyProperties bodyProperties30 = new A.BodyProperties();

            A.ListStyle listStyle30 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties12 = new A.Level1ParagraphProperties();

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing1.Append(spacingPercent1);
            A.DefaultRunProperties defaultRunProperties29 = new A.DefaultRunProperties();

            level1ParagraphProperties12.Append(lineSpacing1);
            level1ParagraphProperties12.Append(defaultRunProperties29);

            listStyle30.Append(level1ParagraphProperties12);

            A.Paragraph paragraph34 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties() { Level = 0 };

            A.Run run31 = new A.Run();
            A.RunProperties runProperties33 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text33 = new A.Text();
            text33.Text = "Click to edit Master text styles";

            run31.Append(runProperties33);
            run31.Append(text33);

            paragraph34.Append(paragraphProperties7);
            paragraph34.Append(run31);

            textBody30.Append(bodyProperties30);
            textBody30.Append(listStyle30);
            textBody30.Append(paragraph34);

            shape35.Append(nonVisualShapeProperties35);
            shape35.Append(shapeProperties35);
            shape35.Append(textBody30);

            Shape shape36 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties36 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties39 = new NonVisualDrawingProperties() { Id = (UInt32Value)12U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties36 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks30 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties36.Append(shapeLocks30);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties39 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape30 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)16U, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties39.Append(placeholderShape30);

            nonVisualShapeProperties36.Append(nonVisualDrawingProperties39);
            nonVisualShapeProperties36.Append(nonVisualShapeDrawingProperties36);
            nonVisualShapeProperties36.Append(applicationNonVisualDrawingProperties39);

            ShapeProperties shapeProperties36 = new ShapeProperties();

            A.Transform2D transform2D25 = new A.Transform2D();
            A.Offset offset28 = new A.Offset() { X = 695325L, Y = 2296391L };
            A.Extents extents28 = new A.Extents() { Cx = 1900238L, Cy = 628578L };

            transform2D25.Append(offset28);
            transform2D25.Append(extents28);

            shapeProperties36.Append(transform2D25);

            TextBody textBody31 = new TextBody();
            A.BodyProperties bodyProperties31 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Top };

            A.ListStyle listStyle31 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties13 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet6 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties30 = new A.DefaultRunProperties() { FontSize = 1400 };
            A.LatinFont latinFont24 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont24 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont24 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties30.Append(latinFont24);
            defaultRunProperties30.Append(eastAsianFont24);
            defaultRunProperties30.Append(complexScriptFont24);

            level1ParagraphProperties13.Append(noBullet6);
            level1ParagraphProperties13.Append(defaultRunProperties30);

            listStyle31.Append(level1ParagraphProperties13);

            A.Paragraph paragraph35 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties() { Level = 0 };

            A.Run run32 = new A.Run();
            A.RunProperties runProperties34 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text34 = new A.Text();
            text34.Text = "CLICK TO EDIT MASTER TEXT STYLES";

            run32.Append(runProperties34);
            run32.Append(text34);

            paragraph35.Append(paragraphProperties8);
            paragraph35.Append(run32);

            textBody31.Append(bodyProperties31);
            textBody31.Append(listStyle31);
            textBody31.Append(paragraph35);

            shape36.Append(nonVisualShapeProperties36);
            shape36.Append(shapeProperties36);
            shape36.Append(textBody31);

            Shape shape37 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties37 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties40 = new NonVisualDrawingProperties() { Id = (UInt32Value)13U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties37 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks31 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties37.Append(shapeLocks31);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties40 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape31 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)17U };

            applicationNonVisualDrawingProperties40.Append(placeholderShape31);

            nonVisualShapeProperties37.Append(nonVisualDrawingProperties40);
            nonVisualShapeProperties37.Append(nonVisualShapeDrawingProperties37);
            nonVisualShapeProperties37.Append(applicationNonVisualDrawingProperties40);

            ShapeProperties shapeProperties37 = new ShapeProperties();

            A.Transform2D transform2D26 = new A.Transform2D();
            A.Offset offset29 = new A.Offset() { X = 695326L, Y = 5116152L };
            A.Extents extents29 = new A.Extents() { Cx = 1900238L, Cy = 1013185L };

            transform2D26.Append(offset29);
            transform2D26.Append(extents29);

            shapeProperties37.Append(transform2D26);

            TextBody textBody32 = new TextBody();
            A.BodyProperties bodyProperties32 = new A.BodyProperties();

            A.ListStyle listStyle32 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties14 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet7 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties31 = new A.DefaultRunProperties() { FontSize = 1200 };
            A.LatinFont latinFont25 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont25 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont25 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties31.Append(latinFont25);
            defaultRunProperties31.Append(eastAsianFont25);
            defaultRunProperties31.Append(complexScriptFont25);

            level1ParagraphProperties14.Append(noBullet7);
            level1ParagraphProperties14.Append(defaultRunProperties31);

            listStyle32.Append(level1ParagraphProperties14);

            A.Paragraph paragraph36 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties() { Level = 0 };

            A.Run run33 = new A.Run();
            A.RunProperties runProperties35 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text35 = new A.Text();
            text35.Text = "Click to edit Master text styles";

            run33.Append(runProperties35);
            run33.Append(text35);

            paragraph36.Append(paragraphProperties9);
            paragraph36.Append(run33);

            textBody32.Append(bodyProperties32);
            textBody32.Append(listStyle32);
            textBody32.Append(paragraph36);

            shape37.Append(nonVisualShapeProperties37);
            shape37.Append(shapeProperties37);
            shape37.Append(textBody32);

            Shape shape38 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties38 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties41 = new NonVisualDrawingProperties() { Id = (UInt32Value)14U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties38 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks32 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties38.Append(shapeLocks32);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties41 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape32 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)18U, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties41.Append(placeholderShape32);

            nonVisualShapeProperties38.Append(nonVisualDrawingProperties41);
            nonVisualShapeProperties38.Append(nonVisualShapeDrawingProperties38);
            nonVisualShapeProperties38.Append(applicationNonVisualDrawingProperties41);

            ShapeProperties shapeProperties38 = new ShapeProperties();

            A.Transform2D transform2D27 = new A.Transform2D();
            A.Offset offset30 = new A.Offset() { X = 2920602L, Y = 2296391L };
            A.Extents extents30 = new A.Extents() { Cx = 1900238L, Cy = 628578L };

            transform2D27.Append(offset30);
            transform2D27.Append(extents30);

            shapeProperties38.Append(transform2D27);

            TextBody textBody33 = new TextBody();
            A.BodyProperties bodyProperties33 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Top };

            A.ListStyle listStyle33 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties15 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet8 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties32 = new A.DefaultRunProperties() { FontSize = 1400 };
            A.LatinFont latinFont26 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont26 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont26 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties32.Append(latinFont26);
            defaultRunProperties32.Append(eastAsianFont26);
            defaultRunProperties32.Append(complexScriptFont26);

            level1ParagraphProperties15.Append(noBullet8);
            level1ParagraphProperties15.Append(defaultRunProperties32);

            listStyle33.Append(level1ParagraphProperties15);

            A.Paragraph paragraph37 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties() { Level = 0 };

            A.Run run34 = new A.Run();
            A.RunProperties runProperties36 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text36 = new A.Text();
            text36.Text = "CLICK TO EDIT MASTER TEXT STYLES";

            run34.Append(runProperties36);
            run34.Append(text36);

            paragraph37.Append(paragraphProperties10);
            paragraph37.Append(run34);

            textBody33.Append(bodyProperties33);
            textBody33.Append(listStyle33);
            textBody33.Append(paragraph37);

            shape38.Append(nonVisualShapeProperties38);
            shape38.Append(shapeProperties38);
            shape38.Append(textBody33);

            Shape shape39 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties39 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties42 = new NonVisualDrawingProperties() { Id = (UInt32Value)15U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties39 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks33 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties39.Append(shapeLocks33);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties42 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape33 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)19U };

            applicationNonVisualDrawingProperties42.Append(placeholderShape33);

            nonVisualShapeProperties39.Append(nonVisualDrawingProperties42);
            nonVisualShapeProperties39.Append(nonVisualShapeDrawingProperties39);
            nonVisualShapeProperties39.Append(applicationNonVisualDrawingProperties42);

            ShapeProperties shapeProperties39 = new ShapeProperties();

            A.Transform2D transform2D28 = new A.Transform2D();
            A.Offset offset31 = new A.Offset() { X = 2920603L, Y = 5116152L };
            A.Extents extents31 = new A.Extents() { Cx = 1900238L, Cy = 1013185L };

            transform2D28.Append(offset31);
            transform2D28.Append(extents31);

            shapeProperties39.Append(transform2D28);

            TextBody textBody34 = new TextBody();
            A.BodyProperties bodyProperties34 = new A.BodyProperties();

            A.ListStyle listStyle34 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties16 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet9 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties33 = new A.DefaultRunProperties() { FontSize = 1200 };
            A.LatinFont latinFont27 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont27 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont27 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties33.Append(latinFont27);
            defaultRunProperties33.Append(eastAsianFont27);
            defaultRunProperties33.Append(complexScriptFont27);

            level1ParagraphProperties16.Append(noBullet9);
            level1ParagraphProperties16.Append(defaultRunProperties33);

            listStyle34.Append(level1ParagraphProperties16);

            A.Paragraph paragraph38 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties() { Level = 0 };

            A.Run run35 = new A.Run();
            A.RunProperties runProperties37 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text37 = new A.Text();
            text37.Text = "Click to edit Master text styles";

            run35.Append(runProperties37);
            run35.Append(text37);

            paragraph38.Append(paragraphProperties11);
            paragraph38.Append(run35);

            textBody34.Append(bodyProperties34);
            textBody34.Append(listStyle34);
            textBody34.Append(paragraph38);

            shape39.Append(nonVisualShapeProperties39);
            shape39.Append(shapeProperties39);
            shape39.Append(textBody34);

            Shape shape40 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties40 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties43 = new NonVisualDrawingProperties() { Id = (UInt32Value)16U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties40 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks34 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties40.Append(shapeLocks34);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties43 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape34 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)20U, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties43.Append(placeholderShape34);

            nonVisualShapeProperties40.Append(nonVisualDrawingProperties43);
            nonVisualShapeProperties40.Append(nonVisualShapeDrawingProperties40);
            nonVisualShapeProperties40.Append(applicationNonVisualDrawingProperties43);

            ShapeProperties shapeProperties40 = new ShapeProperties();

            A.Transform2D transform2D29 = new A.Transform2D();
            A.Offset offset32 = new A.Offset() { X = 5145880L, Y = 2296391L };
            A.Extents extents32 = new A.Extents() { Cx = 1900238L, Cy = 628578L };

            transform2D29.Append(offset32);
            transform2D29.Append(extents32);

            shapeProperties40.Append(transform2D29);

            TextBody textBody35 = new TextBody();
            A.BodyProperties bodyProperties35 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Top };

            A.ListStyle listStyle35 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties17 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet10 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties34 = new A.DefaultRunProperties() { FontSize = 1400 };
            A.LatinFont latinFont28 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont28 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont28 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties34.Append(latinFont28);
            defaultRunProperties34.Append(eastAsianFont28);
            defaultRunProperties34.Append(complexScriptFont28);

            level1ParagraphProperties17.Append(noBullet10);
            level1ParagraphProperties17.Append(defaultRunProperties34);

            listStyle35.Append(level1ParagraphProperties17);

            A.Paragraph paragraph39 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties() { Level = 0 };

            A.Run run36 = new A.Run();
            A.RunProperties runProperties38 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text38 = new A.Text();
            text38.Text = "CLICK TO EDIT MASTER TEXT STYLES";

            run36.Append(runProperties38);
            run36.Append(text38);

            paragraph39.Append(paragraphProperties12);
            paragraph39.Append(run36);

            textBody35.Append(bodyProperties35);
            textBody35.Append(listStyle35);
            textBody35.Append(paragraph39);

            shape40.Append(nonVisualShapeProperties40);
            shape40.Append(shapeProperties40);
            shape40.Append(textBody35);

            Shape shape41 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties41 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties44 = new NonVisualDrawingProperties() { Id = (UInt32Value)17U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks35 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties41.Append(shapeLocks35);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties44 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape35 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)21U };

            applicationNonVisualDrawingProperties44.Append(placeholderShape35);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties44);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties44);

            ShapeProperties shapeProperties41 = new ShapeProperties();

            A.Transform2D transform2D30 = new A.Transform2D();
            A.Offset offset33 = new A.Offset() { X = 5145881L, Y = 5116152L };
            A.Extents extents33 = new A.Extents() { Cx = 1900238L, Cy = 1013185L };

            transform2D30.Append(offset33);
            transform2D30.Append(extents33);

            shapeProperties41.Append(transform2D30);

            TextBody textBody36 = new TextBody();
            A.BodyProperties bodyProperties36 = new A.BodyProperties();

            A.ListStyle listStyle36 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties18 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet11 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties35 = new A.DefaultRunProperties() { FontSize = 1200 };
            A.LatinFont latinFont29 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont29 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont29 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties35.Append(latinFont29);
            defaultRunProperties35.Append(eastAsianFont29);
            defaultRunProperties35.Append(complexScriptFont29);

            level1ParagraphProperties18.Append(noBullet11);
            level1ParagraphProperties18.Append(defaultRunProperties35);

            listStyle36.Append(level1ParagraphProperties18);

            A.Paragraph paragraph40 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties13 = new A.ParagraphProperties() { Level = 0 };

            A.Run run37 = new A.Run();
            A.RunProperties runProperties39 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text39 = new A.Text();
            text39.Text = "Click to edit Master text styles";

            run37.Append(runProperties39);
            run37.Append(text39);

            paragraph40.Append(paragraphProperties13);
            paragraph40.Append(run37);

            textBody36.Append(bodyProperties36);
            textBody36.Append(listStyle36);
            textBody36.Append(paragraph40);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties41);
            shape41.Append(textBody36);

            Shape shape42 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties42 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties45 = new NonVisualDrawingProperties() { Id = (UInt32Value)18U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks36 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties42.Append(shapeLocks36);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties45 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape36 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)22U, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties45.Append(placeholderShape36);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties45);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties45);

            ShapeProperties shapeProperties42 = new ShapeProperties();

            A.Transform2D transform2D31 = new A.Transform2D();
            A.Offset offset34 = new A.Offset() { X = 7371159L, Y = 2296391L };
            A.Extents extents34 = new A.Extents() { Cx = 1900238L, Cy = 628578L };

            transform2D31.Append(offset34);
            transform2D31.Append(extents34);

            shapeProperties42.Append(transform2D31);

            TextBody textBody37 = new TextBody();
            A.BodyProperties bodyProperties37 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Top };

            A.ListStyle listStyle37 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties19 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet12 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties36 = new A.DefaultRunProperties() { FontSize = 1400 };
            A.LatinFont latinFont30 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont30 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont30 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties36.Append(latinFont30);
            defaultRunProperties36.Append(eastAsianFont30);
            defaultRunProperties36.Append(complexScriptFont30);

            level1ParagraphProperties19.Append(noBullet12);
            level1ParagraphProperties19.Append(defaultRunProperties36);

            listStyle37.Append(level1ParagraphProperties19);

            A.Paragraph paragraph41 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties() { Level = 0 };

            A.Run run38 = new A.Run();
            A.RunProperties runProperties40 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text40 = new A.Text();
            text40.Text = "CLICK TO EDIT MASTER TEXT STYLES";

            run38.Append(runProperties40);
            run38.Append(text40);

            paragraph41.Append(paragraphProperties14);
            paragraph41.Append(run38);

            textBody37.Append(bodyProperties37);
            textBody37.Append(listStyle37);
            textBody37.Append(paragraph41);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties42);
            shape42.Append(textBody37);

            Shape shape43 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties43 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties46 = new NonVisualDrawingProperties() { Id = (UInt32Value)19U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks37 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties43.Append(shapeLocks37);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties46 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape37 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)23U };

            applicationNonVisualDrawingProperties46.Append(placeholderShape37);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties46);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties46);

            ShapeProperties shapeProperties43 = new ShapeProperties();

            A.Transform2D transform2D32 = new A.Transform2D();
            A.Offset offset35 = new A.Offset() { X = 7371160L, Y = 5116152L };
            A.Extents extents35 = new A.Extents() { Cx = 1900238L, Cy = 1013185L };

            transform2D32.Append(offset35);
            transform2D32.Append(extents35);

            shapeProperties43.Append(transform2D32);

            TextBody textBody38 = new TextBody();
            A.BodyProperties bodyProperties38 = new A.BodyProperties();

            A.ListStyle listStyle38 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties20 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet13 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties37 = new A.DefaultRunProperties() { FontSize = 1200 };
            A.LatinFont latinFont31 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont31 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont31 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties37.Append(latinFont31);
            defaultRunProperties37.Append(eastAsianFont31);
            defaultRunProperties37.Append(complexScriptFont31);

            level1ParagraphProperties20.Append(noBullet13);
            level1ParagraphProperties20.Append(defaultRunProperties37);

            listStyle38.Append(level1ParagraphProperties20);

            A.Paragraph paragraph42 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties15 = new A.ParagraphProperties() { Level = 0 };

            A.Run run39 = new A.Run();
            A.RunProperties runProperties41 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text41 = new A.Text();
            text41.Text = "Click to edit Master text styles";

            run39.Append(runProperties41);
            run39.Append(text41);

            paragraph42.Append(paragraphProperties15);
            paragraph42.Append(run39);

            textBody38.Append(bodyProperties38);
            textBody38.Append(listStyle38);
            textBody38.Append(paragraph42);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties43);
            shape43.Append(textBody38);

            Shape shape44 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties44 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties47 = new NonVisualDrawingProperties() { Id = (UInt32Value)20U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks38 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties44.Append(shapeLocks38);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties47 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape38 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)24U, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties47.Append(placeholderShape38);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties47);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties47);

            ShapeProperties shapeProperties44 = new ShapeProperties();

            A.Transform2D transform2D33 = new A.Transform2D();
            A.Offset offset36 = new A.Offset() { X = 9596436L, Y = 2296391L };
            A.Extents extents36 = new A.Extents() { Cx = 1900238L, Cy = 628578L };

            transform2D33.Append(offset36);
            transform2D33.Append(extents36);

            shapeProperties44.Append(transform2D33);

            TextBody textBody39 = new TextBody();
            A.BodyProperties bodyProperties39 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Top };

            A.ListStyle listStyle39 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties21 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet14 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties38 = new A.DefaultRunProperties() { FontSize = 1400 };
            A.LatinFont latinFont32 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont32 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont32 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties38.Append(latinFont32);
            defaultRunProperties38.Append(eastAsianFont32);
            defaultRunProperties38.Append(complexScriptFont32);

            level1ParagraphProperties21.Append(noBullet14);
            level1ParagraphProperties21.Append(defaultRunProperties38);

            listStyle39.Append(level1ParagraphProperties21);

            A.Paragraph paragraph43 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties16 = new A.ParagraphProperties() { Level = 0 };

            A.Run run40 = new A.Run();
            A.RunProperties runProperties42 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text42 = new A.Text();
            text42.Text = "CLICK TO EDIT MASTER TEXT STYLES";

            run40.Append(runProperties42);
            run40.Append(text42);

            paragraph43.Append(paragraphProperties16);
            paragraph43.Append(run40);

            textBody39.Append(bodyProperties39);
            textBody39.Append(listStyle39);
            textBody39.Append(paragraph43);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties44);
            shape44.Append(textBody39);

            Shape shape45 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties45 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties48 = new NonVisualDrawingProperties() { Id = (UInt32Value)21U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks39 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties45.Append(shapeLocks39);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties48 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape39 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)25U };

            applicationNonVisualDrawingProperties48.Append(placeholderShape39);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties48);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties48);

            ShapeProperties shapeProperties45 = new ShapeProperties();

            A.Transform2D transform2D34 = new A.Transform2D();
            A.Offset offset37 = new A.Offset() { X = 9596437L, Y = 5116152L };
            A.Extents extents37 = new A.Extents() { Cx = 1900238L, Cy = 1013185L };

            transform2D34.Append(offset37);
            transform2D34.Append(extents37);

            shapeProperties45.Append(transform2D34);

            TextBody textBody40 = new TextBody();
            A.BodyProperties bodyProperties40 = new A.BodyProperties();

            A.ListStyle listStyle40 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties22 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet15 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties39 = new A.DefaultRunProperties() { FontSize = 1200 };
            A.LatinFont latinFont33 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont33 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont33 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties39.Append(latinFont33);
            defaultRunProperties39.Append(eastAsianFont33);
            defaultRunProperties39.Append(complexScriptFont33);

            level1ParagraphProperties22.Append(noBullet15);
            level1ParagraphProperties22.Append(defaultRunProperties39);

            listStyle40.Append(level1ParagraphProperties22);

            A.Paragraph paragraph44 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties17 = new A.ParagraphProperties() { Level = 0 };

            A.Run run41 = new A.Run();
            A.RunProperties runProperties43 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text43 = new A.Text();
            text43.Text = "Click to edit Master text styles";

            run41.Append(runProperties43);
            run41.Append(text43);

            paragraph44.Append(paragraphProperties17);
            paragraph44.Append(run41);

            textBody40.Append(bodyProperties40);
            textBody40.Append(listStyle40);
            textBody40.Append(paragraph44);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties45);
            shape45.Append(textBody40);

            shapeTree3.Append(nonVisualGroupShapeProperties3);
            shapeTree3.Append(groupShapeProperties3);
            shapeTree3.Append(shape29);
            shapeTree3.Append(shape30);
            shapeTree3.Append(shape31);
            shapeTree3.Append(shape32);
            shapeTree3.Append(shape33);
            shapeTree3.Append(shape34);
            shapeTree3.Append(shape35);
            shapeTree3.Append(shape36);
            shapeTree3.Append(shape37);
            shapeTree3.Append(shape38);
            shapeTree3.Append(shape39);
            shapeTree3.Append(shape40);
            shapeTree3.Append(shape41);
            shapeTree3.Append(shape42);
            shapeTree3.Append(shape43);
            shapeTree3.Append(shape44);
            shapeTree3.Append(shape45);

            commonSlideData3.Append(shapeTree3);

            ColorMapOverride colorMapOverride2 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping2 = new A.MasterColorMapping();

            colorMapOverride2.Append(masterColorMapping2);

            slideLayout1.Append(commonSlideData3);
            slideLayout1.Append(colorMapOverride2);

            slideLayoutPart1.SlideLayout = slideLayout1;
        }

        // Generates content of slideMasterPart1.
        private void GenerateSlideMasterPart1Content(SlideMasterPart slideMasterPart1)
        {
            SlideMaster slideMaster1 = new SlideMaster();
            slideMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData4 = new CommonSlideData();

            Background background2 = new Background();

            BackgroundStyleReference backgroundStyleReference2 = new BackgroundStyleReference() { Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference2.Append(schemeColor48);

            background2.Append(backgroundStyleReference2);

            ShapeTree shapeTree4 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties49 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties49 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties49);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties49);

            GroupShapeProperties groupShapeProperties4 = new GroupShapeProperties();

            A.TransformGroup transformGroup4 = new A.TransformGroup();
            A.Offset offset38 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents38 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset4 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents4 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup4.Append(offset38);
            transformGroup4.Append(extents38);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            Shape shape46 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties46 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties50 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks40 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties46.Append(shapeLocks40);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape40 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties50.Append(placeholderShape40);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties50);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties50);

            ShapeProperties shapeProperties46 = new ShapeProperties();

            A.Transform2D transform2D35 = new A.Transform2D();
            A.Offset offset39 = new A.Offset() { X = 695325L, Y = 728663L };
            A.Extents extents39 = new A.Extents() { Cx = 10801350L, Cy = 952193L };

            transform2D35.Append(offset39);
            transform2D35.Append(extents39);

            A.PresetGeometry presetGeometry18 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList18 = new A.AdjustValueList();

            presetGeometry18.Append(adjustValueList18);

            shapeProperties46.Append(transform2D35);
            shapeProperties46.Append(presetGeometry18);

            TextBody textBody41 = new TextBody();

            A.BodyProperties bodyProperties41 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Top };
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            bodyProperties41.Append(noAutoFit1);
            A.ListStyle listStyle41 = new A.ListStyle();

            A.Paragraph paragraph45 = new A.Paragraph();

            A.Run run42 = new A.Run();
            A.RunProperties runProperties44 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text44 = new A.Text();
            text44.Text = "Click to edit Master title style";

            run42.Append(runProperties44);
            run42.Append(text44);

            paragraph45.Append(run42);

            textBody41.Append(bodyProperties41);
            textBody41.Append(listStyle41);
            textBody41.Append(paragraph45);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties46);
            shape46.Append(textBody41);

            Shape shape47 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties47 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties51 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks41 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties47.Append(shapeLocks41);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape41 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties51.Append(placeholderShape41);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties51);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties51);

            ShapeProperties shapeProperties47 = new ShapeProperties();

            A.Transform2D transform2D36 = new A.Transform2D();
            A.Offset offset40 = new A.Offset() { X = 695325L, Y = 1680857L };
            A.Extents extents40 = new A.Extents() { Cx = 10801350L, Cy = 4448482L };

            transform2D36.Append(offset40);
            transform2D36.Append(extents40);

            A.PresetGeometry presetGeometry19 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList19 = new A.AdjustValueList();

            presetGeometry19.Append(adjustValueList19);

            shapeProperties47.Append(transform2D36);
            shapeProperties47.Append(presetGeometry19);

            TextBody textBody42 = new TextBody();

            A.BodyProperties bodyProperties42 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

            bodyProperties42.Append(noAutoFit2);
            A.ListStyle listStyle42 = new A.ListStyle();

            A.Paragraph paragraph46 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties18 = new A.ParagraphProperties() { Level = 0 };

            A.Run run43 = new A.Run();
            A.RunProperties runProperties45 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text45 = new A.Text();
            text45.Text = "Click to edit Master text styles";

            run43.Append(runProperties45);
            run43.Append(text45);

            paragraph46.Append(paragraphProperties18);
            paragraph46.Append(run43);

            A.Paragraph paragraph47 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties19 = new A.ParagraphProperties() { Level = 1 };

            A.Run run44 = new A.Run();
            A.RunProperties runProperties46 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text46 = new A.Text();
            text46.Text = "Second level";

            run44.Append(runProperties46);
            run44.Append(text46);

            paragraph47.Append(paragraphProperties19);
            paragraph47.Append(run44);

            A.Paragraph paragraph48 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties20 = new A.ParagraphProperties() { Level = 2 };

            A.Run run45 = new A.Run();
            A.RunProperties runProperties47 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text47 = new A.Text();
            text47.Text = "Third level";

            run45.Append(runProperties47);
            run45.Append(text47);

            paragraph48.Append(paragraphProperties20);
            paragraph48.Append(run45);

            textBody42.Append(bodyProperties42);
            textBody42.Append(listStyle42);
            textBody42.Append(paragraph46);
            textBody42.Append(paragraph47);
            textBody42.Append(paragraph48);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties47);
            shape47.Append(textBody42);

            Picture picture1 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties52 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties52);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties52);

            BlipFill blipFill1 = new BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId11" };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties48 = new ShapeProperties();

            A.Transform2D transform2D37 = new A.Transform2D();
            A.Offset offset41 = new A.Offset() { X = 10919356L, Y = 6465900L };
            A.Extents extents41 = new A.Extents() { Cx = 1095427L, Cy = 260968L };

            transform2D37.Append(offset41);
            transform2D37.Append(extents41);

            A.PresetGeometry presetGeometry20 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList20 = new A.AdjustValueList();

            presetGeometry20.Append(adjustValueList20);

            shapeProperties48.Append(transform2D37);
            shapeProperties48.Append(presetGeometry20);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties48);

            shapeTree4.Append(nonVisualGroupShapeProperties4);
            shapeTree4.Append(groupShapeProperties4);
            shapeTree4.Append(shape46);
            shapeTree4.Append(shape47);
            shapeTree4.Append(picture1);

            CommonSlideDataExtensionList commonSlideDataExtensionList3 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension3 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId3 = new P14.CreationId() { Val = (UInt32Value)35696427U };
            creationId3.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension3.Append(creationId3);

            commonSlideDataExtensionList3.Append(commonSlideDataExtension3);

            commonSlideData4.Append(background2);
            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList3);
            ColorMap colorMap2 = new ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            SlideLayoutIdList slideLayoutIdList1 = new SlideLayoutIdList();
            SlideLayoutId slideLayoutId1 = new SlideLayoutId() { Id = (UInt32Value)2147483681U, RelationshipId = "rId1" };
            SlideLayoutId slideLayoutId2 = new SlideLayoutId() { Id = (UInt32Value)2147483682U, RelationshipId = "rId2" };
            SlideLayoutId slideLayoutId3 = new SlideLayoutId() { Id = (UInt32Value)2147483683U, RelationshipId = "rId3" };
            SlideLayoutId slideLayoutId4 = new SlideLayoutId() { Id = (UInt32Value)2147483684U, RelationshipId = "rId4" };
            SlideLayoutId slideLayoutId5 = new SlideLayoutId() { Id = (UInt32Value)2147483685U, RelationshipId = "rId5" };
            SlideLayoutId slideLayoutId6 = new SlideLayoutId() { Id = (UInt32Value)2147483686U, RelationshipId = "rId6" };
            SlideLayoutId slideLayoutId7 = new SlideLayoutId() { Id = (UInt32Value)2147483687U, RelationshipId = "rId7" };
            SlideLayoutId slideLayoutId8 = new SlideLayoutId() { Id = (UInt32Value)2147483679U, RelationshipId = "rId8" };
            SlideLayoutId slideLayoutId9 = new SlideLayoutId() { Id = (UInt32Value)2147483680U, RelationshipId = "rId9" };

            slideLayoutIdList1.Append(slideLayoutId1);
            slideLayoutIdList1.Append(slideLayoutId2);
            slideLayoutIdList1.Append(slideLayoutId3);
            slideLayoutIdList1.Append(slideLayoutId4);
            slideLayoutIdList1.Append(slideLayoutId5);
            slideLayoutIdList1.Append(slideLayoutId6);
            slideLayoutIdList1.Append(slideLayoutId7);
            slideLayoutIdList1.Append(slideLayoutId8);
            slideLayoutIdList1.Append(slideLayoutId9);

            TextStyles textStyles1 = new TextStyles();

            TitleStyle titleStyle1 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties23 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing2.Append(spacingPercent2);

            A.SpaceBefore spaceBefore2 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent3 = new A.SpacingPercent() { Val = 0 };

            spaceBefore2.Append(spacingPercent3);
            A.NoBullet noBullet16 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties40 = new A.DefaultRunProperties() { FontSize = 3600, Bold = true, Kerning = 1200, Spacing = -150 };

            A.SolidFill solidFill36 = new A.SolidFill();
            A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill36.Append(schemeColor49);
            A.LatinFont latinFont34 = new A.LatinFont() { Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont34 = new A.EastAsianFont() { Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont34 = new A.ComplexScriptFont() { Typeface = "+mj-cs" };

            defaultRunProperties40.Append(solidFill36);
            defaultRunProperties40.Append(latinFont34);
            defaultRunProperties40.Append(eastAsianFont34);
            defaultRunProperties40.Append(complexScriptFont34);

            level1ParagraphProperties23.Append(lineSpacing2);
            level1ParagraphProperties23.Append(spaceBefore2);
            level1ParagraphProperties23.Append(noBullet16);
            level1ParagraphProperties23.Append(defaultRunProperties40);

            titleStyle1.Append(level1ParagraphProperties23);

            BodyStyle bodyStyle1 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties24 = new A.Level1ParagraphProperties() { LeftMargin = 228600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing3 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing3.Append(spacingPercent4);

            A.SpaceBefore spaceBefore3 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints2 = new A.SpacingPoints() { Val = 1000 };

            spaceBefore3.Append(spacingPoints2);
            A.BulletFont bulletFont1 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties41 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill37 = new A.SolidFill();
            A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill37.Append(schemeColor50);
            A.LatinFont latinFont35 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont35 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont35 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties41.Append(solidFill37);
            defaultRunProperties41.Append(latinFont35);
            defaultRunProperties41.Append(eastAsianFont35);
            defaultRunProperties41.Append(complexScriptFont35);

            level1ParagraphProperties24.Append(lineSpacing3);
            level1ParagraphProperties24.Append(spaceBefore3);
            level1ParagraphProperties24.Append(bulletFont1);
            level1ParagraphProperties24.Append(characterBullet1);
            level1ParagraphProperties24.Append(defaultRunProperties41);

            A.Level2ParagraphProperties level2ParagraphProperties3 = new A.Level2ParagraphProperties() { LeftMargin = 495300, Indent = -227013, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing4 = new A.LineSpacing();
            A.SpacingPercent spacingPercent5 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing4.Append(spacingPercent5);

            A.SpaceBefore spaceBefore4 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints3 = new A.SpacingPoints() { Val = 500 };

            spaceBefore4.Append(spacingPoints3);
            A.BulletFont bulletFont2 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet() { Char = "•" };
            A.TabStopList tabStopList1 = new A.TabStopList();

            A.DefaultRunProperties defaultRunProperties42 = new A.DefaultRunProperties() { FontSize = 1600, Kerning = 1200 };

            A.SolidFill solidFill38 = new A.SolidFill();
            A.SchemeColor schemeColor51 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill38.Append(schemeColor51);
            A.LatinFont latinFont36 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont36 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont36 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties42.Append(solidFill38);
            defaultRunProperties42.Append(latinFont36);
            defaultRunProperties42.Append(eastAsianFont36);
            defaultRunProperties42.Append(complexScriptFont36);

            level2ParagraphProperties3.Append(lineSpacing4);
            level2ParagraphProperties3.Append(spaceBefore4);
            level2ParagraphProperties3.Append(bulletFont2);
            level2ParagraphProperties3.Append(characterBullet2);
            level2ParagraphProperties3.Append(tabStopList1);
            level2ParagraphProperties3.Append(defaultRunProperties42);

            A.Level3ParagraphProperties level3ParagraphProperties3 = new A.Level3ParagraphProperties() { LeftMargin = 711200, Indent = -215900, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing5 = new A.LineSpacing();
            A.SpacingPercent spacingPercent6 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing5.Append(spacingPercent6);

            A.SpaceBefore spaceBefore5 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints4 = new A.SpacingPoints() { Val = 500 };

            spaceBefore5.Append(spacingPoints4);
            A.BulletFont bulletFont3 = new A.BulletFont() { Typeface = ".AppleSystemUIFont", CharacterSet = -120 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet() { Char = "-" };
            A.TabStopList tabStopList2 = new A.TabStopList();

            A.DefaultRunProperties defaultRunProperties43 = new A.DefaultRunProperties() { FontSize = 1400, Kerning = 1200 };

            A.SolidFill solidFill39 = new A.SolidFill();
            A.SchemeColor schemeColor52 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill39.Append(schemeColor52);
            A.LatinFont latinFont37 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont37 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont37 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties43.Append(solidFill39);
            defaultRunProperties43.Append(latinFont37);
            defaultRunProperties43.Append(eastAsianFont37);
            defaultRunProperties43.Append(complexScriptFont37);

            level3ParagraphProperties3.Append(lineSpacing5);
            level3ParagraphProperties3.Append(spaceBefore5);
            level3ParagraphProperties3.Append(bulletFont3);
            level3ParagraphProperties3.Append(characterBullet3);
            level3ParagraphProperties3.Append(tabStopList2);
            level3ParagraphProperties3.Append(defaultRunProperties43);

            A.Level4ParagraphProperties level4ParagraphProperties3 = new A.Level4ParagraphProperties() { LeftMargin = 1600200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing6 = new A.LineSpacing();
            A.SpacingPercent spacingPercent7 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing6.Append(spacingPercent7);

            A.SpaceBefore spaceBefore6 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints5 = new A.SpacingPoints() { Val = 500 };

            spaceBefore6.Append(spacingPoints5);
            A.BulletFont bulletFont4 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties44 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill40 = new A.SolidFill();
            A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill40.Append(schemeColor53);
            A.LatinFont latinFont38 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont38 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont38 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties44.Append(solidFill40);
            defaultRunProperties44.Append(latinFont38);
            defaultRunProperties44.Append(eastAsianFont38);
            defaultRunProperties44.Append(complexScriptFont38);

            level4ParagraphProperties3.Append(lineSpacing6);
            level4ParagraphProperties3.Append(spaceBefore6);
            level4ParagraphProperties3.Append(bulletFont4);
            level4ParagraphProperties3.Append(characterBullet4);
            level4ParagraphProperties3.Append(defaultRunProperties44);

            A.Level5ParagraphProperties level5ParagraphProperties3 = new A.Level5ParagraphProperties() { LeftMargin = 2057400, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing7 = new A.LineSpacing();
            A.SpacingPercent spacingPercent8 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing7.Append(spacingPercent8);

            A.SpaceBefore spaceBefore7 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints6 = new A.SpacingPoints() { Val = 500 };

            spaceBefore7.Append(spacingPoints6);
            A.BulletFont bulletFont5 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties45 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill41 = new A.SolidFill();
            A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill41.Append(schemeColor54);
            A.LatinFont latinFont39 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont39 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont39 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties45.Append(solidFill41);
            defaultRunProperties45.Append(latinFont39);
            defaultRunProperties45.Append(eastAsianFont39);
            defaultRunProperties45.Append(complexScriptFont39);

            level5ParagraphProperties3.Append(lineSpacing7);
            level5ParagraphProperties3.Append(spaceBefore7);
            level5ParagraphProperties3.Append(bulletFont5);
            level5ParagraphProperties3.Append(characterBullet5);
            level5ParagraphProperties3.Append(defaultRunProperties45);

            A.Level6ParagraphProperties level6ParagraphProperties3 = new A.Level6ParagraphProperties() { LeftMargin = 2514600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing8 = new A.LineSpacing();
            A.SpacingPercent spacingPercent9 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing8.Append(spacingPercent9);

            A.SpaceBefore spaceBefore8 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints7 = new A.SpacingPoints() { Val = 500 };

            spaceBefore8.Append(spacingPoints7);
            A.BulletFont bulletFont6 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties46 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill42 = new A.SolidFill();
            A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill42.Append(schemeColor55);
            A.LatinFont latinFont40 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont40 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont40 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties46.Append(solidFill42);
            defaultRunProperties46.Append(latinFont40);
            defaultRunProperties46.Append(eastAsianFont40);
            defaultRunProperties46.Append(complexScriptFont40);

            level6ParagraphProperties3.Append(lineSpacing8);
            level6ParagraphProperties3.Append(spaceBefore8);
            level6ParagraphProperties3.Append(bulletFont6);
            level6ParagraphProperties3.Append(characterBullet6);
            level6ParagraphProperties3.Append(defaultRunProperties46);

            A.Level7ParagraphProperties level7ParagraphProperties3 = new A.Level7ParagraphProperties() { LeftMargin = 2971800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing9 = new A.LineSpacing();
            A.SpacingPercent spacingPercent10 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing9.Append(spacingPercent10);

            A.SpaceBefore spaceBefore9 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints8 = new A.SpacingPoints() { Val = 500 };

            spaceBefore9.Append(spacingPoints8);
            A.BulletFont bulletFont7 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties47 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill43 = new A.SolidFill();
            A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill43.Append(schemeColor56);
            A.LatinFont latinFont41 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont41 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont41 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties47.Append(solidFill43);
            defaultRunProperties47.Append(latinFont41);
            defaultRunProperties47.Append(eastAsianFont41);
            defaultRunProperties47.Append(complexScriptFont41);

            level7ParagraphProperties3.Append(lineSpacing9);
            level7ParagraphProperties3.Append(spaceBefore9);
            level7ParagraphProperties3.Append(bulletFont7);
            level7ParagraphProperties3.Append(characterBullet7);
            level7ParagraphProperties3.Append(defaultRunProperties47);

            A.Level8ParagraphProperties level8ParagraphProperties3 = new A.Level8ParagraphProperties() { LeftMargin = 3429000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing10 = new A.LineSpacing();
            A.SpacingPercent spacingPercent11 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing10.Append(spacingPercent11);

            A.SpaceBefore spaceBefore10 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints9 = new A.SpacingPoints() { Val = 500 };

            spaceBefore10.Append(spacingPoints9);
            A.BulletFont bulletFont8 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties48 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill44 = new A.SolidFill();
            A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill44.Append(schemeColor57);
            A.LatinFont latinFont42 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont42 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont42 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties48.Append(solidFill44);
            defaultRunProperties48.Append(latinFont42);
            defaultRunProperties48.Append(eastAsianFont42);
            defaultRunProperties48.Append(complexScriptFont42);

            level8ParagraphProperties3.Append(lineSpacing10);
            level8ParagraphProperties3.Append(spaceBefore10);
            level8ParagraphProperties3.Append(bulletFont8);
            level8ParagraphProperties3.Append(characterBullet8);
            level8ParagraphProperties3.Append(defaultRunProperties48);

            A.Level9ParagraphProperties level9ParagraphProperties3 = new A.Level9ParagraphProperties() { LeftMargin = 3886200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing11 = new A.LineSpacing();
            A.SpacingPercent spacingPercent12 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing11.Append(spacingPercent12);

            A.SpaceBefore spaceBefore11 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints10 = new A.SpacingPoints() { Val = 500 };

            spaceBefore11.Append(spacingPoints10);
            A.BulletFont bulletFont9 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties49 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill45 = new A.SolidFill();
            A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill45.Append(schemeColor58);
            A.LatinFont latinFont43 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont43 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont43 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties49.Append(solidFill45);
            defaultRunProperties49.Append(latinFont43);
            defaultRunProperties49.Append(eastAsianFont43);
            defaultRunProperties49.Append(complexScriptFont43);

            level9ParagraphProperties3.Append(lineSpacing11);
            level9ParagraphProperties3.Append(spaceBefore11);
            level9ParagraphProperties3.Append(bulletFont9);
            level9ParagraphProperties3.Append(characterBullet9);
            level9ParagraphProperties3.Append(defaultRunProperties49);

            bodyStyle1.Append(level1ParagraphProperties24);
            bodyStyle1.Append(level2ParagraphProperties3);
            bodyStyle1.Append(level3ParagraphProperties3);
            bodyStyle1.Append(level4ParagraphProperties3);
            bodyStyle1.Append(level5ParagraphProperties3);
            bodyStyle1.Append(level6ParagraphProperties3);
            bodyStyle1.Append(level7ParagraphProperties3);
            bodyStyle1.Append(level8ParagraphProperties3);
            bodyStyle1.Append(level9ParagraphProperties3);

            OtherStyle otherStyle1 = new OtherStyle();

            A.DefaultParagraphProperties defaultParagraphProperties2 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties50 = new A.DefaultRunProperties() { Language = "en-US" };

            defaultParagraphProperties2.Append(defaultRunProperties50);

            A.Level1ParagraphProperties level1ParagraphProperties25 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties51 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill46 = new A.SolidFill();
            A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill46.Append(schemeColor59);
            A.LatinFont latinFont44 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont44 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont44 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties51.Append(solidFill46);
            defaultRunProperties51.Append(latinFont44);
            defaultRunProperties51.Append(eastAsianFont44);
            defaultRunProperties51.Append(complexScriptFont44);

            level1ParagraphProperties25.Append(defaultRunProperties51);

            A.Level2ParagraphProperties level2ParagraphProperties4 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties52 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill47 = new A.SolidFill();
            A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill47.Append(schemeColor60);
            A.LatinFont latinFont45 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont45 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont45 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties52.Append(solidFill47);
            defaultRunProperties52.Append(latinFont45);
            defaultRunProperties52.Append(eastAsianFont45);
            defaultRunProperties52.Append(complexScriptFont45);

            level2ParagraphProperties4.Append(defaultRunProperties52);

            A.Level3ParagraphProperties level3ParagraphProperties4 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties53 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill48 = new A.SolidFill();
            A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill48.Append(schemeColor61);
            A.LatinFont latinFont46 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont46 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont46 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties53.Append(solidFill48);
            defaultRunProperties53.Append(latinFont46);
            defaultRunProperties53.Append(eastAsianFont46);
            defaultRunProperties53.Append(complexScriptFont46);

            level3ParagraphProperties4.Append(defaultRunProperties53);

            A.Level4ParagraphProperties level4ParagraphProperties4 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties54 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill49 = new A.SolidFill();
            A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill49.Append(schemeColor62);
            A.LatinFont latinFont47 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont47 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont47 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties54.Append(solidFill49);
            defaultRunProperties54.Append(latinFont47);
            defaultRunProperties54.Append(eastAsianFont47);
            defaultRunProperties54.Append(complexScriptFont47);

            level4ParagraphProperties4.Append(defaultRunProperties54);

            A.Level5ParagraphProperties level5ParagraphProperties4 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties55 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill50 = new A.SolidFill();
            A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill50.Append(schemeColor63);
            A.LatinFont latinFont48 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont48 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont48 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties55.Append(solidFill50);
            defaultRunProperties55.Append(latinFont48);
            defaultRunProperties55.Append(eastAsianFont48);
            defaultRunProperties55.Append(complexScriptFont48);

            level5ParagraphProperties4.Append(defaultRunProperties55);

            A.Level6ParagraphProperties level6ParagraphProperties4 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties56 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill51 = new A.SolidFill();
            A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill51.Append(schemeColor64);
            A.LatinFont latinFont49 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont49 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont49 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties56.Append(solidFill51);
            defaultRunProperties56.Append(latinFont49);
            defaultRunProperties56.Append(eastAsianFont49);
            defaultRunProperties56.Append(complexScriptFont49);

            level6ParagraphProperties4.Append(defaultRunProperties56);

            A.Level7ParagraphProperties level7ParagraphProperties4 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties57 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill52 = new A.SolidFill();
            A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill52.Append(schemeColor65);
            A.LatinFont latinFont50 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont50 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont50 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties57.Append(solidFill52);
            defaultRunProperties57.Append(latinFont50);
            defaultRunProperties57.Append(eastAsianFont50);
            defaultRunProperties57.Append(complexScriptFont50);

            level7ParagraphProperties4.Append(defaultRunProperties57);

            A.Level8ParagraphProperties level8ParagraphProperties4 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties58 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill53 = new A.SolidFill();
            A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill53.Append(schemeColor66);
            A.LatinFont latinFont51 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont51 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont51 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties58.Append(solidFill53);
            defaultRunProperties58.Append(latinFont51);
            defaultRunProperties58.Append(eastAsianFont51);
            defaultRunProperties58.Append(complexScriptFont51);

            level8ParagraphProperties4.Append(defaultRunProperties58);

            A.Level9ParagraphProperties level9ParagraphProperties4 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties59 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill54 = new A.SolidFill();
            A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill54.Append(schemeColor67);
            A.LatinFont latinFont52 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont52 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont52 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties59.Append(solidFill54);
            defaultRunProperties59.Append(latinFont52);
            defaultRunProperties59.Append(eastAsianFont52);
            defaultRunProperties59.Append(complexScriptFont52);

            level9ParagraphProperties4.Append(defaultRunProperties59);

            otherStyle1.Append(defaultParagraphProperties2);
            otherStyle1.Append(level1ParagraphProperties25);
            otherStyle1.Append(level2ParagraphProperties4);
            otherStyle1.Append(level3ParagraphProperties4);
            otherStyle1.Append(level4ParagraphProperties4);
            otherStyle1.Append(level5ParagraphProperties4);
            otherStyle1.Append(level6ParagraphProperties4);
            otherStyle1.Append(level7ParagraphProperties4);
            otherStyle1.Append(level8ParagraphProperties4);
            otherStyle1.Append(level9ParagraphProperties4);

            textStyles1.Append(titleStyle1);
            textStyles1.Append(bodyStyle1);
            textStyles1.Append(otherStyle1);

            SlideMasterExtensionList slideMasterExtensionList1 = new SlideMasterExtensionList();
            slideMasterExtensionList1.SetAttribute(new OpenXmlAttribute("", "mod", "", "1"));

            SlideMasterExtension slideMasterExtension1 = new SlideMasterExtension() { Uri = "{27BBF7A9-308A-43DC-89C8-2F10F3537804}" };

            P15.SlideGuideList slideGuideList2 = new P15.SlideGuideList();
            slideGuideList2.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            P15.ExtendedGuide extendedGuide3 = new P15.ExtendedGuide() { Id = (UInt32Value)1U, Orientation = DirectionValues.Horizontal, Position = 2160, IsUserDrawn = true };

            P15.ColorType colorType3 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex15 = new A.RgbColorModelHex() { Val = "F26B43" };

            colorType3.Append(rgbColorModelHex15);

            extendedGuide3.Append(colorType3);

            P15.ExtendedGuide extendedGuide4 = new P15.ExtendedGuide() { Id = (UInt32Value)2U, Position = 3840, IsUserDrawn = true };

            P15.ColorType colorType4 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex16 = new A.RgbColorModelHex() { Val = "F26B43" };

            colorType4.Append(rgbColorModelHex16);

            extendedGuide4.Append(colorType4);

            P15.ExtendedGuide extendedGuide5 = new P15.ExtendedGuide() { Id = (UInt32Value)3U, Position = 438, IsUserDrawn = true };

            P15.ColorType colorType5 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex17 = new A.RgbColorModelHex() { Val = "F26B43" };

            colorType5.Append(rgbColorModelHex17);

            extendedGuide5.Append(colorType5);

            P15.ExtendedGuide extendedGuide6 = new P15.ExtendedGuide() { Id = (UInt32Value)4U, Position = 7242, IsUserDrawn = true };

            P15.ColorType colorType6 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex18 = new A.RgbColorModelHex() { Val = "F26B43" };

            colorType6.Append(rgbColorModelHex18);

            extendedGuide6.Append(colorType6);

            P15.ExtendedGuide extendedGuide7 = new P15.ExtendedGuide() { Id = (UInt32Value)5U, Orientation = DirectionValues.Horizontal, Position = 459, IsUserDrawn = true };

            P15.ColorType colorType7 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex19 = new A.RgbColorModelHex() { Val = "F26B43" };

            colorType7.Append(rgbColorModelHex19);

            extendedGuide7.Append(colorType7);

            P15.ExtendedGuide extendedGuide8 = new P15.ExtendedGuide() { Id = (UInt32Value)6U, Orientation = DirectionValues.Horizontal, Position = 3861, IsUserDrawn = true };

            P15.ColorType colorType8 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex20 = new A.RgbColorModelHex() { Val = "F26B43" };

            colorType8.Append(rgbColorModelHex20);

            extendedGuide8.Append(colorType8);

            P15.ExtendedGuide extendedGuide9 = new P15.ExtendedGuide() { Id = (UInt32Value)7U, Orientation = DirectionValues.Horizontal, Position = 2441, IsUserDrawn = true };

            P15.ColorType colorType9 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex21 = new A.RgbColorModelHex() { Val = "F26B43" };

            colorType9.Append(rgbColorModelHex21);

            extendedGuide9.Append(colorType9);

            slideGuideList2.Append(extendedGuide3);
            slideGuideList2.Append(extendedGuide4);
            slideGuideList2.Append(extendedGuide5);
            slideGuideList2.Append(extendedGuide6);
            slideGuideList2.Append(extendedGuide7);
            slideGuideList2.Append(extendedGuide8);
            slideGuideList2.Append(extendedGuide9);

            slideMasterExtension1.Append(slideGuideList2);

            slideMasterExtensionList1.Append(slideMasterExtension1);

            slideMaster1.Append(commonSlideData4);
            slideMaster1.Append(colorMap2);
            slideMaster1.Append(slideLayoutIdList1);
            slideMaster1.Append(textStyles1);
            slideMaster1.Append(slideMasterExtensionList1);

            slideMasterPart1.SlideMaster = slideMaster1;
        }

        // Generates content of slideLayoutPart2.
        private void GenerateSlideLayoutPart2Content(SlideLayoutPart slideLayoutPart2)
        {
            SlideLayout slideLayout2 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData5 = new CommonSlideData() { Name = "Main Menu" };

            ShapeTree shapeTree5 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties5 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties53 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties5 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties5.Append(nonVisualDrawingProperties53);
            nonVisualGroupShapeProperties5.Append(nonVisualGroupShapeDrawingProperties5);
            nonVisualGroupShapeProperties5.Append(applicationNonVisualDrawingProperties53);

            GroupShapeProperties groupShapeProperties5 = new GroupShapeProperties();

            A.TransformGroup transformGroup5 = new A.TransformGroup();
            A.Offset offset42 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents42 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset5 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents5 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup5.Append(offset42);
            transformGroup5.Append(extents42);
            transformGroup5.Append(childOffset5);
            transformGroup5.Append(childExtents5);

            groupShapeProperties5.Append(transformGroup5);

            Picture picture2 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties2 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties54 = new NonVisualDrawingProperties() { Id = (UInt32Value)11U, Name = "Picture 10" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties2 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks2 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties2.Append(pictureLocks2);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties2.Append(nonVisualDrawingProperties54);
            nonVisualPictureProperties2.Append(nonVisualPictureDrawingProperties2);
            nonVisualPictureProperties2.Append(applicationNonVisualDrawingProperties54);

            BlipFill blipFill2 = new BlipFill();

            A.Blip blip2 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList2 = new A.BlipExtensionList();

            A.BlipExtension blipExtension2 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi2 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi2.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension2.Append(useLocalDpi2);

            blipExtensionList2.Append(blipExtension2);

            blip2.Append(blipExtensionList2);

            A.Stretch stretch2 = new A.Stretch();
            A.FillRectangle fillRectangle2 = new A.FillRectangle();

            stretch2.Append(fillRectangle2);

            blipFill2.Append(blip2);
            blipFill2.Append(stretch2);

            ShapeProperties shapeProperties49 = new ShapeProperties();

            A.Transform2D transform2D38 = new A.Transform2D();
            A.Offset offset43 = new A.Offset() { X = -16772L, Y = 0L };
            A.Extents extents43 = new A.Extents() { Cx = 12208772L, Cy = 6858000L };

            transform2D38.Append(offset43);
            transform2D38.Append(extents43);

            A.PresetGeometry presetGeometry21 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList21 = new A.AdjustValueList();

            presetGeometry21.Append(adjustValueList21);

            shapeProperties49.Append(transform2D38);
            shapeProperties49.Append(presetGeometry21);

            picture2.Append(nonVisualPictureProperties2);
            picture2.Append(blipFill2);
            picture2.Append(shapeProperties49);

            Shape shape48 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties48 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties55 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Rectangle 1" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties55);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties55);

            ShapeProperties shapeProperties50 = new ShapeProperties();

            A.Transform2D transform2D39 = new A.Transform2D();
            A.Offset offset44 = new A.Offset() { X = -16772L, Y = 0L };
            A.Extents extents44 = new A.Extents() { Cx = 12208772L, Cy = 451411L };

            transform2D39.Append(offset44);
            transform2D39.Append(extents44);

            A.PresetGeometry presetGeometry22 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList22 = new A.AdjustValueList();

            presetGeometry22.Append(adjustValueList22);

            A.SolidFill solidFill55 = new A.SolidFill();
            A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill55.Append(schemeColor68);

            A.Outline outline6 = new A.Outline();
            A.NoFill noFill8 = new A.NoFill();

            outline6.Append(noFill8);

            shapeProperties50.Append(transform2D39);
            shapeProperties50.Append(presetGeometry22);
            shapeProperties50.Append(solidFill55);
            shapeProperties50.Append(outline6);

            ShapeStyle shapeStyle2 = new ShapeStyle();

            A.LineReference lineReference2 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade7 = new A.Shade() { Val = 50000 };

            schemeColor69.Append(shade7);

            lineReference2.Append(schemeColor69);

            A.FillReference fillReference2 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference2.Append(schemeColor70);

            A.EffectReference effectReference2 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference2.Append(schemeColor71);

            A.FontReference fontReference2 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference2.Append(schemeColor72);

            shapeStyle2.Append(lineReference2);
            shapeStyle2.Append(fillReference2);
            shapeStyle2.Append(effectReference2);
            shapeStyle2.Append(fontReference2);

            TextBody textBody43 = new TextBody();
            A.BodyProperties bodyProperties43 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle43 = new A.ListStyle();

            A.Paragraph paragraph49 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties21 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties18 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph49.Append(paragraphProperties21);
            paragraph49.Append(endParagraphRunProperties18);

            textBody43.Append(bodyProperties43);
            textBody43.Append(listStyle43);
            textBody43.Append(paragraph49);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties50);
            shape48.Append(shapeStyle2);
            shape48.Append(textBody43);

            Picture picture3 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties3 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties56 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture 2" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties3 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks3 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties3.Append(pictureLocks3);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties3.Append(nonVisualDrawingProperties56);
            nonVisualPictureProperties3.Append(nonVisualPictureDrawingProperties3);
            nonVisualPictureProperties3.Append(applicationNonVisualDrawingProperties56);

            BlipFill blipFill3 = new BlipFill();

            A.Blip blip3 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList3 = new A.BlipExtensionList();

            A.BlipExtension blipExtension3 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi3 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi3.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension3.Append(useLocalDpi3);

            blipExtensionList3.Append(blipExtension3);

            blip3.Append(blipExtensionList3);

            A.Stretch stretch3 = new A.Stretch();
            A.FillRectangle fillRectangle3 = new A.FillRectangle();

            stretch3.Append(fillRectangle3);

            blipFill3.Append(blip3);
            blipFill3.Append(stretch3);

            ShapeProperties shapeProperties51 = new ShapeProperties();

            A.Transform2D transform2D40 = new A.Transform2D();
            A.Offset offset45 = new A.Offset() { X = 237321L, Y = 83863L };
            A.Extents extents45 = new A.Extents() { Cx = 1190776L, Cy = 283684L };

            transform2D40.Append(offset45);
            transform2D40.Append(extents45);

            A.PresetGeometry presetGeometry23 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList23 = new A.AdjustValueList();

            presetGeometry23.Append(adjustValueList23);
            A.NoFill noFill9 = new A.NoFill();

            A.Outline outline7 = new A.Outline();
            A.NoFill noFill10 = new A.NoFill();

            outline7.Append(noFill10);

            shapeProperties51.Append(transform2D40);
            shapeProperties51.Append(presetGeometry23);
            shapeProperties51.Append(noFill9);
            shapeProperties51.Append(outline7);

            picture3.Append(nonVisualPictureProperties3);
            picture3.Append(blipFill3);
            picture3.Append(shapeProperties51);

            Picture picture4 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties4 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties57 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties4 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks4 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties4.Append(pictureLocks4);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties4.Append(nonVisualDrawingProperties57);
            nonVisualPictureProperties4.Append(nonVisualPictureDrawingProperties4);
            nonVisualPictureProperties4.Append(applicationNonVisualDrawingProperties57);

            BlipFill blipFill4 = new BlipFill() { RotateWithShape = true };

            A.Blip blip4 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList4 = new A.BlipExtensionList();

            A.BlipExtension blipExtension4 = new A.BlipExtension() { Uri = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}" };

            A14.ImageProperties imageProperties1 = new A14.ImageProperties();
            imageProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A14.ImageLayer imageLayer1 = new A14.ImageLayer() { Embed = "rId3" };

            A14.ImageEffect imageEffect1 = new A14.ImageEffect();
            A14.BrightnessContrast brightnessContrast1 = new A14.BrightnessContrast() { Bright = 100000 };

            imageEffect1.Append(brightnessContrast1);

            imageLayer1.Append(imageEffect1);

            imageProperties1.Append(imageLayer1);

            blipExtension4.Append(imageProperties1);

            blipExtensionList4.Append(blipExtension4);

            blip4.Append(blipExtensionList4);
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle() { Left = -36828, Top = -36828, Right = -36828, Bottom = -36828 };
            A.Stretch stretch4 = new A.Stretch();

            blipFill4.Append(blip4);
            blipFill4.Append(sourceRectangle1);
            blipFill4.Append(stretch4);

            ShapeProperties shapeProperties52 = new ShapeProperties();

            A.Transform2D transform2D41 = new A.Transform2D();
            A.Offset offset46 = new A.Offset() { X = 5978873L, Y = 100485L };
            A.Extents extents46 = new A.Extents() { Cx = 267062L, Cy = 267062L };

            transform2D41.Append(offset46);
            transform2D41.Append(extents46);

            A.PresetGeometry presetGeometry24 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList24 = new A.AdjustValueList();

            presetGeometry24.Append(adjustValueList24);

            A.SolidFill solidFill56 = new A.SolidFill();
            A.SchemeColor schemeColor73 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill56.Append(schemeColor73);

            shapeProperties52.Append(transform2D41);
            shapeProperties52.Append(presetGeometry24);
            shapeProperties52.Append(solidFill56);

            picture4.Append(nonVisualPictureProperties4);
            picture4.Append(blipFill4);
            picture4.Append(shapeProperties52);

            Shape shape49 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties49 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties58 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "TextBox 4" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties49 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualShapeProperties49.Append(nonVisualDrawingProperties58);
            nonVisualShapeProperties49.Append(nonVisualShapeDrawingProperties49);
            nonVisualShapeProperties49.Append(applicationNonVisualDrawingProperties58);

            ShapeProperties shapeProperties53 = new ShapeProperties();

            A.Transform2D transform2D42 = new A.Transform2D();
            A.Offset offset47 = new A.Offset() { X = 8594830L, Y = 109217L };
            A.Extents extents47 = new A.Extents() { Cx = 3042208L, Cy = 276999L };

            transform2D42.Append(offset47);
            transform2D42.Append(extents47);

            A.PresetGeometry presetGeometry25 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList25 = new A.AdjustValueList();

            presetGeometry25.Append(adjustValueList25);
            A.NoFill noFill11 = new A.NoFill();

            shapeProperties53.Append(transform2D42);
            shapeProperties53.Append(presetGeometry25);
            shapeProperties53.Append(noFill11);

            TextBody textBody44 = new TextBody();

            A.BodyProperties bodyProperties44 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit6 = new A.ShapeAutoFit();

            bodyProperties44.Append(shapeAutoFit6);
            A.ListStyle listStyle44 = new A.ListStyle();

            A.Paragraph paragraph50 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties22 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Right };

            A.Run run46 = new A.Run();

            A.RunProperties runProperties48 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Dirty = false };
            A.LatinFont latinFont53 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont53 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont53 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties48.Append(latinFont53);
            runProperties48.Append(eastAsianFont53);
            runProperties48.Append(complexScriptFont53);
            A.Text text48 = new A.Text();
            text48.Text = "More about ";

            run46.Append(runProperties48);
            run46.Append(text48);

            A.Run run47 = new A.Run();

            A.RunProperties runProperties49 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Dirty = false, SpellingError = true };
            A.LatinFont latinFont54 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont54 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont54 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties49.Append(latinFont54);
            runProperties49.Append(eastAsianFont54);
            runProperties49.Append(complexScriptFont54);
            A.Text text49 = new A.Text();
            text49.Text = "Discover.ai";

            run47.Append(runProperties49);
            run47.Append(text49);

            A.EndParagraphRunProperties endParagraphRunProperties19 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Dirty = false };
            A.LatinFont latinFont55 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont55 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont55 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties19.Append(latinFont55);
            endParagraphRunProperties19.Append(eastAsianFont55);
            endParagraphRunProperties19.Append(complexScriptFont55);

            paragraph50.Append(paragraphProperties22);
            paragraph50.Append(run46);
            paragraph50.Append(run47);
            paragraph50.Append(endParagraphRunProperties19);

            textBody44.Append(bodyProperties44);
            textBody44.Append(listStyle44);
            textBody44.Append(paragraph50);

            shape49.Append(nonVisualShapeProperties49);
            shape49.Append(shapeProperties53);
            shape49.Append(textBody44);

            GroupShape groupShape1 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties6 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties59 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Group 5" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties6 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties59 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualGroupShapeProperties6.Append(nonVisualDrawingProperties59);
            nonVisualGroupShapeProperties6.Append(nonVisualGroupShapeDrawingProperties6);
            nonVisualGroupShapeProperties6.Append(applicationNonVisualDrawingProperties59);

            GroupShapeProperties groupShapeProperties6 = new GroupShapeProperties();

            A.TransformGroup transformGroup6 = new A.TransformGroup();
            A.Offset offset48 = new A.Offset() { X = 11637038L, Y = 100485L };
            A.Extents extents48 = new A.Extents() { Cx = 267062L, Cy = 266400L };
            A.ChildOffset childOffset6 = new A.ChildOffset() { X = 10356783L, Y = 2945331L };
            A.ChildExtents childExtents6 = new A.ChildExtents() { Cx = 1139892L, Cy = 1139892L };

            transformGroup6.Append(offset48);
            transformGroup6.Append(extents48);
            transformGroup6.Append(childOffset6);
            transformGroup6.Append(childExtents6);

            groupShapeProperties6.Append(transformGroup6);

            Shape shape50 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties50 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties60 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Oval 6" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties50 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties60 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties50.Append(nonVisualDrawingProperties60);
            nonVisualShapeProperties50.Append(nonVisualShapeDrawingProperties50);
            nonVisualShapeProperties50.Append(applicationNonVisualDrawingProperties60);

            ShapeProperties shapeProperties54 = new ShapeProperties();

            A.Transform2D transform2D43 = new A.Transform2D();
            A.Offset offset49 = new A.Offset() { X = 10356783L, Y = 2945331L };
            A.Extents extents49 = new A.Extents() { Cx = 1139892L, Cy = 1139892L };

            transform2D43.Append(offset49);
            transform2D43.Append(extents49);

            A.PresetGeometry presetGeometry26 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList26 = new A.AdjustValueList();

            presetGeometry26.Append(adjustValueList26);

            A.SolidFill solidFill57 = new A.SolidFill();
            A.SchemeColor schemeColor74 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill57.Append(schemeColor74);

            A.Outline outline8 = new A.Outline();
            A.NoFill noFill12 = new A.NoFill();

            outline8.Append(noFill12);

            shapeProperties54.Append(transform2D43);
            shapeProperties54.Append(presetGeometry26);
            shapeProperties54.Append(solidFill57);
            shapeProperties54.Append(outline8);

            ShapeStyle shapeStyle3 = new ShapeStyle();

            A.LineReference lineReference3 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor75 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade8 = new A.Shade() { Val = 50000 };

            schemeColor75.Append(shade8);

            lineReference3.Append(schemeColor75);

            A.FillReference fillReference3 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor76 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference3.Append(schemeColor76);

            A.EffectReference effectReference3 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor77 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference3.Append(schemeColor77);

            A.FontReference fontReference3 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor78 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference3.Append(schemeColor78);

            shapeStyle3.Append(lineReference3);
            shapeStyle3.Append(fillReference3);
            shapeStyle3.Append(effectReference3);
            shapeStyle3.Append(fontReference3);

            TextBody textBody45 = new TextBody();
            A.BodyProperties bodyProperties45 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle45 = new A.ListStyle();

            A.Paragraph paragraph51 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties23 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties20 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph51.Append(paragraphProperties23);
            paragraph51.Append(endParagraphRunProperties20);

            textBody45.Append(bodyProperties45);
            textBody45.Append(listStyle45);
            textBody45.Append(paragraph51);

            shape50.Append(nonVisualShapeProperties50);
            shape50.Append(shapeProperties54);
            shape50.Append(shapeStyle3);
            shape50.Append(textBody45);

            GroupShape groupShape2 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties7 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties61 = new NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Group 7" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties7 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties61 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties7.Append(nonVisualDrawingProperties61);
            nonVisualGroupShapeProperties7.Append(nonVisualGroupShapeDrawingProperties7);
            nonVisualGroupShapeProperties7.Append(applicationNonVisualDrawingProperties61);

            GroupShapeProperties groupShapeProperties7 = new GroupShapeProperties();

            A.TransformGroup transformGroup7 = new A.TransformGroup();
            A.Offset offset50 = new A.Offset() { X = 10804795L, Y = 3324431L };
            A.Extents extents50 = new A.Extents() { Cx = 243867L, Cy = 381692L };
            A.ChildOffset childOffset7 = new A.ChildOffset() { X = 6996222L, Y = 4465674L };
            A.ChildExtents childExtents7 = new A.ChildExtents() { Cx = 122276L, Cy = 191386L };

            transformGroup7.Append(offset50);
            transformGroup7.Append(extents50);
            transformGroup7.Append(childOffset7);
            transformGroup7.Append(childExtents7);

            groupShapeProperties7.Append(transformGroup7);

            ConnectionShape connectionShape1 = new ConnectionShape();

            NonVisualConnectionShapeProperties nonVisualConnectionShapeProperties1 = new NonVisualConnectionShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties62 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Straight Connector 8" };
            NonVisualConnectorShapeDrawingProperties nonVisualConnectorShapeDrawingProperties1 = new NonVisualConnectorShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties62 = new ApplicationNonVisualDrawingProperties();

            nonVisualConnectionShapeProperties1.Append(nonVisualDrawingProperties62);
            nonVisualConnectionShapeProperties1.Append(nonVisualConnectorShapeDrawingProperties1);
            nonVisualConnectionShapeProperties1.Append(applicationNonVisualDrawingProperties62);

            ShapeProperties shapeProperties55 = new ShapeProperties();

            A.Transform2D transform2D44 = new A.Transform2D();
            A.Offset offset51 = new A.Offset() { X = 6996222L, Y = 4465674L };
            A.Extents extents51 = new A.Extents() { Cx = 122276L, Cy = 92685L };

            transform2D44.Append(offset51);
            transform2D44.Append(extents51);

            A.PresetGeometry presetGeometry27 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Line };
            A.AdjustValueList adjustValueList27 = new A.AdjustValueList();

            presetGeometry27.Append(adjustValueList27);

            A.Outline outline9 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill58 = new A.SolidFill();
            A.SchemeColor schemeColor79 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill58.Append(schemeColor79);
            A.Round round1 = new A.Round();

            outline9.Append(solidFill58);
            outline9.Append(round1);

            shapeProperties55.Append(transform2D44);
            shapeProperties55.Append(presetGeometry27);
            shapeProperties55.Append(outline9);

            ShapeStyle shapeStyle4 = new ShapeStyle();

            A.LineReference lineReference4 = new A.LineReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor80 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference4.Append(schemeColor80);

            A.FillReference fillReference4 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor81 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference4.Append(schemeColor81);

            A.EffectReference effectReference4 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor82 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference4.Append(schemeColor82);

            A.FontReference fontReference4 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor83 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference4.Append(schemeColor83);

            shapeStyle4.Append(lineReference4);
            shapeStyle4.Append(fillReference4);
            shapeStyle4.Append(effectReference4);
            shapeStyle4.Append(fontReference4);

            connectionShape1.Append(nonVisualConnectionShapeProperties1);
            connectionShape1.Append(shapeProperties55);
            connectionShape1.Append(shapeStyle4);

            ConnectionShape connectionShape2 = new ConnectionShape();

            NonVisualConnectionShapeProperties nonVisualConnectionShapeProperties2 = new NonVisualConnectionShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties63 = new NonVisualDrawingProperties() { Id = (UInt32Value)10U, Name = "Straight Connector 9" };
            NonVisualConnectorShapeDrawingProperties nonVisualConnectorShapeDrawingProperties2 = new NonVisualConnectorShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties63 = new ApplicationNonVisualDrawingProperties();

            nonVisualConnectionShapeProperties2.Append(nonVisualDrawingProperties63);
            nonVisualConnectionShapeProperties2.Append(nonVisualConnectorShapeDrawingProperties2);
            nonVisualConnectionShapeProperties2.Append(applicationNonVisualDrawingProperties63);

            ShapeProperties shapeProperties56 = new ShapeProperties();

            A.Transform2D transform2D45 = new A.Transform2D() { HorizontalFlip = true };
            A.Offset offset52 = new A.Offset() { X = 6996222L, Y = 4561367L };
            A.Extents extents52 = new A.Extents() { Cx = 122276L, Cy = 95693L };

            transform2D45.Append(offset52);
            transform2D45.Append(extents52);

            A.PresetGeometry presetGeometry28 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Line };
            A.AdjustValueList adjustValueList28 = new A.AdjustValueList();

            presetGeometry28.Append(adjustValueList28);

            A.Outline outline10 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill59 = new A.SolidFill();
            A.SchemeColor schemeColor84 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill59.Append(schemeColor84);
            A.Round round2 = new A.Round();

            outline10.Append(solidFill59);
            outline10.Append(round2);

            shapeProperties56.Append(transform2D45);
            shapeProperties56.Append(presetGeometry28);
            shapeProperties56.Append(outline10);

            ShapeStyle shapeStyle5 = new ShapeStyle();

            A.LineReference lineReference5 = new A.LineReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor85 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference5.Append(schemeColor85);

            A.FillReference fillReference5 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor86 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference5.Append(schemeColor86);

            A.EffectReference effectReference5 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor87 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference5.Append(schemeColor87);

            A.FontReference fontReference5 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor88 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference5.Append(schemeColor88);

            shapeStyle5.Append(lineReference5);
            shapeStyle5.Append(fillReference5);
            shapeStyle5.Append(effectReference5);
            shapeStyle5.Append(fontReference5);

            connectionShape2.Append(nonVisualConnectionShapeProperties2);
            connectionShape2.Append(shapeProperties56);
            connectionShape2.Append(shapeStyle5);

            groupShape2.Append(nonVisualGroupShapeProperties7);
            groupShape2.Append(groupShapeProperties7);
            groupShape2.Append(connectionShape1);
            groupShape2.Append(connectionShape2);

            groupShape1.Append(nonVisualGroupShapeProperties6);
            groupShape1.Append(groupShapeProperties6);
            groupShape1.Append(shape50);
            groupShape1.Append(groupShape2);

            Shape shape51 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties51 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties64 = new NonVisualDrawingProperties() { Id = (UInt32Value)13U, Name = "Text Placeholder 12" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties51 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks42 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties51.Append(shapeLocks42);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties64 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape42 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties64.Append(placeholderShape42);

            nonVisualShapeProperties51.Append(nonVisualDrawingProperties64);
            nonVisualShapeProperties51.Append(nonVisualShapeDrawingProperties51);
            nonVisualShapeProperties51.Append(applicationNonVisualDrawingProperties64);

            ShapeProperties shapeProperties57 = new ShapeProperties();

            A.Transform2D transform2D46 = new A.Transform2D();
            A.Offset offset53 = new A.Offset() { X = -17463L, Y = 450850L };
            A.Extents extents53 = new A.Extents() { Cx = 12209463L, Cy = 6407150L };

            transform2D46.Append(offset53);
            transform2D46.Append(extents53);

            A.SolidFill solidFill60 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex22 = new A.RgbColorModelHex() { Val = "86D5AC" };
            A.Alpha alpha2 = new A.Alpha() { Val = 80000 };

            rgbColorModelHex22.Append(alpha2);

            solidFill60.Append(rgbColorModelHex22);

            shapeProperties57.Append(transform2D46);
            shapeProperties57.Append(solidFill60);

            TextBody textBody46 = new TextBody();

            A.BodyProperties bodyProperties46 = new A.BodyProperties();
            A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit();

            bodyProperties46.Append(normalAutoFit1);

            A.ListStyle listStyle46 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties26 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet17 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties60 = new A.DefaultRunProperties() { FontSize = 1400 };

            A.SolidFill solidFill61 = new A.SolidFill();

            A.SchemeColor schemeColor89 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.Alpha alpha3 = new A.Alpha() { Val = 0 };

            schemeColor89.Append(alpha3);

            solidFill61.Append(schemeColor89);

            defaultRunProperties60.Append(solidFill61);

            level1ParagraphProperties26.Append(noBullet17);
            level1ParagraphProperties26.Append(defaultRunProperties60);

            listStyle46.Append(level1ParagraphProperties26);

            A.Paragraph paragraph52 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties24 = new A.ParagraphProperties() { Level = 0 };

            A.Run run48 = new A.Run();
            A.RunProperties runProperties50 = new A.RunProperties() { Language = "en-US" };
            A.Text text50 = new A.Text();
            text50.Text = "Click to edit Master text styles";

            run48.Append(runProperties50);
            run48.Append(text50);

            paragraph52.Append(paragraphProperties24);
            paragraph52.Append(run48);

            textBody46.Append(bodyProperties46);
            textBody46.Append(listStyle46);
            textBody46.Append(paragraph52);

            shape51.Append(nonVisualShapeProperties51);
            shape51.Append(shapeProperties57);
            shape51.Append(textBody46);

            Shape shape52 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties52 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties65 = new NonVisualDrawingProperties() { Id = (UInt32Value)14U, Name = "Rectangle 13" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties52 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualShapeProperties52.Append(nonVisualDrawingProperties65);
            nonVisualShapeProperties52.Append(nonVisualShapeDrawingProperties52);
            nonVisualShapeProperties52.Append(applicationNonVisualDrawingProperties65);

            ShapeProperties shapeProperties58 = new ShapeProperties();

            A.Transform2D transform2D47 = new A.Transform2D();
            A.Offset offset54 = new A.Offset() { X = 6112404L, Y = 451412L };
            A.Extents extents54 = new A.Extents() { Cx = 6079596L, Cy = 3429000L };

            transform2D47.Append(offset54);
            transform2D47.Append(extents54);

            A.PresetGeometry presetGeometry29 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList29 = new A.AdjustValueList();

            presetGeometry29.Append(adjustValueList29);

            A.SolidFill solidFill62 = new A.SolidFill();
            A.SchemeColor schemeColor90 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill62.Append(schemeColor90);

            A.Outline outline11 = new A.Outline();
            A.NoFill noFill13 = new A.NoFill();

            outline11.Append(noFill13);

            shapeProperties58.Append(transform2D47);
            shapeProperties58.Append(presetGeometry29);
            shapeProperties58.Append(solidFill62);
            shapeProperties58.Append(outline11);

            ShapeStyle shapeStyle6 = new ShapeStyle();

            A.LineReference lineReference6 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor91 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade9 = new A.Shade() { Val = 50000 };

            schemeColor91.Append(shade9);

            lineReference6.Append(schemeColor91);

            A.FillReference fillReference6 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor92 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference6.Append(schemeColor92);

            A.EffectReference effectReference6 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor93 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference6.Append(schemeColor93);

            A.FontReference fontReference6 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor94 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference6.Append(schemeColor94);

            shapeStyle6.Append(lineReference6);
            shapeStyle6.Append(fillReference6);
            shapeStyle6.Append(effectReference6);
            shapeStyle6.Append(fontReference6);

            TextBody textBody47 = new TextBody();
            A.BodyProperties bodyProperties47 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle47 = new A.ListStyle();

            A.Paragraph paragraph53 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties25 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.Run run49 = new A.Run();

            A.RunProperties runProperties51 = new A.RunProperties() { Language = "en-US", FontSize = 1600, Dirty = false };
            A.LatinFont latinFont56 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont56 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont56 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties51.Append(latinFont56);
            runProperties51.Append(eastAsianFont56);
            runProperties51.Append(complexScriptFont56);
            A.Text text51 = new A.Text();
            text51.Text = "[AN INTRODUCTION TO ";

            run49.Append(runProperties51);
            run49.Append(text51);

            A.Break break1 = new A.Break();

            A.RunProperties runProperties52 = new A.RunProperties() { Language = "en-US", FontSize = 1600, Dirty = false };
            A.LatinFont latinFont57 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont57 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont57 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties52.Append(latinFont57);
            runProperties52.Append(eastAsianFont57);
            runProperties52.Append(complexScriptFont57);

            break1.Append(runProperties52);

            A.Run run50 = new A.Run();

            A.RunProperties runProperties53 = new A.RunProperties() { Language = "en-US", FontSize = 1600, Dirty = false, SpellingError = true };
            A.LatinFont latinFont58 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont58 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont58 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties53.Append(latinFont58);
            runProperties53.Append(eastAsianFont58);
            runProperties53.Append(complexScriptFont58);
            A.Text text52 = new A.Text();
            text52.Text = "DISCOVER.AI";

            run50.Append(runProperties53);
            run50.Append(text52);

            A.Run run51 = new A.Run();

            A.RunProperties runProperties54 = new A.RunProperties() { Language = "en-US", FontSize = 1600, Dirty = false };
            A.LatinFont latinFont59 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont59 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont59 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties54.Append(latinFont59);
            runProperties54.Append(eastAsianFont59);
            runProperties54.Append(complexScriptFont59);
            A.Text text53 = new A.Text();
            text53.Text = " VIDEO]";

            run51.Append(runProperties54);
            run51.Append(text53);

            paragraph53.Append(paragraphProperties25);
            paragraph53.Append(run49);
            paragraph53.Append(break1);
            paragraph53.Append(run50);
            paragraph53.Append(run51);

            textBody47.Append(bodyProperties47);
            textBody47.Append(listStyle47);
            textBody47.Append(paragraph53);

            shape52.Append(nonVisualShapeProperties52);
            shape52.Append(shapeProperties58);
            shape52.Append(shapeStyle6);
            shape52.Append(textBody47);

            Shape shape53 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties53 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties66 = new NonVisualDrawingProperties() { Id = (UInt32Value)16U, Name = "Media Placeholder 15" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties53 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks43 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties53.Append(shapeLocks43);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape43 = new PlaceholderShape() { Type = PlaceholderValues.Media, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties66.Append(placeholderShape43);

            nonVisualShapeProperties53.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties53.Append(nonVisualShapeDrawingProperties53);
            nonVisualShapeProperties53.Append(applicationNonVisualDrawingProperties66);

            ShapeProperties shapeProperties59 = new ShapeProperties();

            A.Transform2D transform2D48 = new A.Transform2D();
            A.Offset offset55 = new A.Offset() { X = 6112404L, Y = 450850L };
            A.Extents extents55 = new A.Extents() { Cx = 6079596L, Cy = 3430588L };

            transform2D48.Append(offset55);
            transform2D48.Append(extents55);

            A.SolidFill solidFill63 = new A.SolidFill();
            A.SchemeColor schemeColor95 = new A.SchemeColor() { Val = A.SchemeColorValues.Background2 };

            solidFill63.Append(schemeColor95);

            shapeProperties59.Append(transform2D48);
            shapeProperties59.Append(solidFill63);

            TextBody textBody48 = new TextBody();

            A.BodyProperties bodyProperties48 = new A.BodyProperties();
            A.NormalAutoFit normalAutoFit2 = new A.NormalAutoFit();

            bodyProperties48.Append(normalAutoFit2);

            A.ListStyle listStyle48 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties27 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet18 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties61 = new A.DefaultRunProperties() { FontSize = 1800 };

            level1ParagraphProperties27.Append(noBullet18);
            level1ParagraphProperties27.Append(defaultRunProperties61);

            listStyle48.Append(level1ParagraphProperties27);

            A.Paragraph paragraph54 = new A.Paragraph();

            A.Run run52 = new A.Run();
            A.RunProperties runProperties55 = new A.RunProperties() { Language = "en-US" };
            A.Text text54 = new A.Text();
            text54.Text = "Insert video";

            run52.Append(runProperties55);
            run52.Append(text54);

            paragraph54.Append(run52);

            textBody48.Append(bodyProperties48);
            textBody48.Append(listStyle48);
            textBody48.Append(paragraph54);

            shape53.Append(nonVisualShapeProperties53);
            shape53.Append(shapeProperties59);
            shape53.Append(textBody48);

            Shape shape54 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties54 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties67 = new NonVisualDrawingProperties() { Id = (UInt32Value)18U, Name = "Picture Placeholder 17" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks44 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties54.Append(shapeLocks44);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape44 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties67.Append(placeholderShape44);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties67);

            ShapeProperties shapeProperties60 = new ShapeProperties();

            A.Transform2D transform2D49 = new A.Transform2D();
            A.Offset offset56 = new A.Offset() { X = -17463L, Y = 450850L };
            A.Extents extents56 = new A.Extents() { Cx = 6129338L, Cy = 3429000L };

            transform2D49.Append(offset56);
            transform2D49.Append(extents56);
            A.NoFill noFill14 = new A.NoFill();

            shapeProperties60.Append(transform2D49);
            shapeProperties60.Append(noFill14);

            TextBody textBody49 = new TextBody();

            A.BodyProperties bodyProperties49 = new A.BodyProperties();
            A.NormalAutoFit normalAutoFit3 = new A.NormalAutoFit();

            bodyProperties49.Append(normalAutoFit3);

            A.ListStyle listStyle49 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties28 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.BulletFont bulletFont10 = new A.BulletFont() { Typeface = "Arial", CharacterSet = 0 };
            A.NoBullet noBullet19 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties62 = new A.DefaultRunProperties() { FontSize = 2000 };

            level1ParagraphProperties28.Append(bulletFont10);
            level1ParagraphProperties28.Append(noBullet19);
            level1ParagraphProperties28.Append(defaultRunProperties62);

            listStyle49.Append(level1ParagraphProperties28);

            A.Paragraph paragraph55 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties21 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph55.Append(endParagraphRunProperties21);

            textBody49.Append(bodyProperties49);
            textBody49.Append(listStyle49);
            textBody49.Append(paragraph55);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties60);
            shape54.Append(textBody49);

            Shape shape55 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties55 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties68 = new NonVisualDrawingProperties() { Id = (UInt32Value)22U, Name = "Title 21" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks45 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties55.Append(shapeLocks45);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape45 = new PlaceholderShape() { Type = PlaceholderValues.Title, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties68.Append(placeholderShape45);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties68);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties68);

            ShapeProperties shapeProperties61 = new ShapeProperties();

            A.Transform2D transform2D50 = new A.Transform2D();
            A.Offset offset57 = new A.Offset() { X = 567224L, Y = 1207811L };
            A.Extents extents57 = new A.Extents() { Cx = 4961184L, Cy = 1346200L };

            transform2D50.Append(offset57);
            transform2D50.Append(extents57);

            shapeProperties61.Append(transform2D50);

            TextBody textBody50 = new TextBody();

            A.BodyProperties bodyProperties50 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };
            A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

            bodyProperties50.Append(noAutoFit3);

            A.ListStyle listStyle50 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties29 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties63 = new A.DefaultRunProperties() { FontSize = 8000, Bold = true, Spacing = -300 };

            A.SolidFill solidFill64 = new A.SolidFill();
            A.SchemeColor schemeColor96 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill64.Append(schemeColor96);

            defaultRunProperties63.Append(solidFill64);

            level1ParagraphProperties29.Append(defaultRunProperties63);

            listStyle50.Append(level1ParagraphProperties29);

            A.Paragraph paragraph56 = new A.Paragraph();

            A.Run run53 = new A.Run();
            A.RunProperties runProperties56 = new A.RunProperties() { Language = "en-US" };
            A.Text text55 = new A.Text();
            text55.Text = "Title";

            run53.Append(runProperties56);
            run53.Append(text55);
            A.EndParagraphRunProperties endParagraphRunProperties22 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph56.Append(run53);
            paragraph56.Append(endParagraphRunProperties22);

            textBody50.Append(bodyProperties50);
            textBody50.Append(listStyle50);
            textBody50.Append(paragraph56);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties61);
            shape55.Append(textBody50);

            Shape shape56 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties56 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties69 = new NonVisualDrawingProperties() { Id = (UInt32Value)24U, Name = "Text Placeholder 23" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks46 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties56.Append(shapeLocks46);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape46 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)13U };

            applicationNonVisualDrawingProperties69.Append(placeholderShape46);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties69);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties69);

            ShapeProperties shapeProperties62 = new ShapeProperties();

            A.Transform2D transform2D51 = new A.Transform2D();
            A.Offset offset58 = new A.Offset() { X = 566349L, Y = 2445610L };
            A.Extents extents58 = new A.Extents() { Cx = 4961713L, Cy = 558800L };

            transform2D51.Append(offset58);
            transform2D51.Append(extents58);

            shapeProperties62.Append(transform2D51);

            TextBody textBody51 = new TextBody();

            A.BodyProperties bodyProperties51 = new A.BodyProperties();
            A.NoAutoFit noAutoFit4 = new A.NoAutoFit();

            bodyProperties51.Append(noAutoFit4);

            A.ListStyle listStyle51 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties30 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet20 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties64 = new A.DefaultRunProperties() { FontSize = 1400 };

            A.SolidFill solidFill65 = new A.SolidFill();
            A.SchemeColor schemeColor97 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill65.Append(schemeColor97);
            A.LatinFont latinFont60 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont60 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont60 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties64.Append(solidFill65);
            defaultRunProperties64.Append(latinFont60);
            defaultRunProperties64.Append(eastAsianFont60);
            defaultRunProperties64.Append(complexScriptFont60);

            level1ParagraphProperties30.Append(noBullet20);
            level1ParagraphProperties30.Append(defaultRunProperties64);

            listStyle51.Append(level1ParagraphProperties30);

            A.Paragraph paragraph57 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties26 = new A.ParagraphProperties() { Level = 0 };

            A.Run run54 = new A.Run();
            A.RunProperties runProperties57 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text56 = new A.Text();
            text56.Text = "Click to edit Master ";

            run54.Append(runProperties57);
            run54.Append(text56);

            A.Run run55 = new A.Run();
            A.RunProperties runProperties58 = new A.RunProperties() { Language = "en-US" };
            A.Text text57 = new A.Text();
            text57.Text = "text styles";

            run55.Append(runProperties58);
            run55.Append(text57);
            A.EndParagraphRunProperties endParagraphRunProperties23 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph57.Append(paragraphProperties26);
            paragraph57.Append(run54);
            paragraph57.Append(run55);
            paragraph57.Append(endParagraphRunProperties23);

            textBody51.Append(bodyProperties51);
            textBody51.Append(listStyle51);
            textBody51.Append(paragraph57);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties62);
            shape56.Append(textBody51);

            Shape shape57 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties57 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties70 = new NonVisualDrawingProperties() { Id = (UInt32Value)26U, Name = "Text Placeholder 23" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks47 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties57.Append(shapeLocks47);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape47 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)14U, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties70.Append(placeholderShape47);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties70);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties70);

            ShapeProperties shapeProperties63 = new ShapeProperties();

            A.Transform2D transform2D52 = new A.Transform2D();
            A.Offset offset59 = new A.Offset() { X = 3633117L, Y = 3951174L };
            A.Extents extents59 = new A.Extents() { Cx = 4961713L, Cy = 291035L };

            transform2D52.Append(offset59);
            transform2D52.Append(extents59);

            shapeProperties63.Append(transform2D52);

            TextBody textBody52 = new TextBody();

            A.BodyProperties bodyProperties52 = new A.BodyProperties();
            A.NoAutoFit noAutoFit5 = new A.NoAutoFit();

            bodyProperties52.Append(noAutoFit5);

            A.ListStyle listStyle52 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties31 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet21 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties65 = new A.DefaultRunProperties() { FontSize = 1200 };

            A.SolidFill solidFill66 = new A.SolidFill();
            A.SchemeColor schemeColor98 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill66.Append(schemeColor98);
            A.LatinFont latinFont61 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont61 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont61 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties65.Append(solidFill66);
            defaultRunProperties65.Append(latinFont61);
            defaultRunProperties65.Append(eastAsianFont61);
            defaultRunProperties65.Append(complexScriptFont61);

            level1ParagraphProperties31.Append(noBullet21);
            level1ParagraphProperties31.Append(defaultRunProperties65);

            listStyle52.Append(level1ParagraphProperties31);

            A.Paragraph paragraph58 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties27 = new A.ParagraphProperties() { Level = 0 };

            A.Run run56 = new A.Run();
            A.RunProperties runProperties59 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text58 = new A.Text();
            text58.Text = "CLICK TO EDIT MASTER TEXT STYLES";

            run56.Append(runProperties59);
            run56.Append(text58);

            paragraph58.Append(paragraphProperties27);
            paragraph58.Append(run56);

            textBody52.Append(bodyProperties52);
            textBody52.Append(listStyle52);
            textBody52.Append(paragraph58);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties63);
            shape57.Append(textBody52);

            shapeTree5.Append(nonVisualGroupShapeProperties5);
            shapeTree5.Append(groupShapeProperties5);
            shapeTree5.Append(picture2);
            shapeTree5.Append(shape48);
            shapeTree5.Append(picture3);
            shapeTree5.Append(picture4);
            shapeTree5.Append(shape49);
            shapeTree5.Append(groupShape1);
            shapeTree5.Append(shape51);
            shapeTree5.Append(shape52);
            shapeTree5.Append(shape53);
            shapeTree5.Append(shape54);
            shapeTree5.Append(shape55);
            shapeTree5.Append(shape56);
            shapeTree5.Append(shape57);

            commonSlideData5.Append(shapeTree5);

            ColorMapOverride colorMapOverride3 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout2.Append(commonSlideData5);
            slideLayout2.Append(colorMapOverride3);

            slideLayoutPart2.SlideLayout = slideLayout2;
        }

        // Generates content of slideLayoutPart3.
        private void GenerateSlideLayoutPart3Content(SlideLayoutPart slideLayoutPart3)
        {
            SlideLayout slideLayout3 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData6 = new CommonSlideData() { Name = "Title Content" };

            ShapeTree shapeTree6 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties8 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties71 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties8 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties8.Append(nonVisualDrawingProperties71);
            nonVisualGroupShapeProperties8.Append(nonVisualGroupShapeDrawingProperties8);
            nonVisualGroupShapeProperties8.Append(applicationNonVisualDrawingProperties71);

            GroupShapeProperties groupShapeProperties8 = new GroupShapeProperties();

            A.TransformGroup transformGroup8 = new A.TransformGroup();
            A.Offset offset60 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents60 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset8 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents8 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup8.Append(offset60);
            transformGroup8.Append(extents60);
            transformGroup8.Append(childOffset8);
            transformGroup8.Append(childExtents8);

            groupShapeProperties8.Append(transformGroup8);

            Shape shape58 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties58 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties72 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks48 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties58.Append(shapeLocks48);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties72 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape48 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties72.Append(placeholderShape48);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties72);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties72);
            ShapeProperties shapeProperties64 = new ShapeProperties();

            TextBody textBody53 = new TextBody();
            A.BodyProperties bodyProperties53 = new A.BodyProperties();
            A.ListStyle listStyle53 = new A.ListStyle();

            A.Paragraph paragraph59 = new A.Paragraph();

            A.Run run57 = new A.Run();
            A.RunProperties runProperties60 = new A.RunProperties() { Language = "en-US" };
            A.Text text59 = new A.Text();
            text59.Text = "Click to edit Master title style";

            run57.Append(runProperties60);
            run57.Append(text59);

            paragraph59.Append(run57);

            textBody53.Append(bodyProperties53);
            textBody53.Append(listStyle53);
            textBody53.Append(paragraph59);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties64);
            shape58.Append(textBody53);

            Shape shape59 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties59 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties73 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks49 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties59.Append(shapeLocks49);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties73 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape49 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties73.Append(placeholderShape49);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties73);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties73);
            ShapeProperties shapeProperties65 = new ShapeProperties();

            TextBody textBody54 = new TextBody();
            A.BodyProperties bodyProperties54 = new A.BodyProperties();

            A.ListStyle listStyle54 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties32 = new A.Level1ParagraphProperties();

            A.LineSpacing lineSpacing12 = new A.LineSpacing();
            A.SpacingPercent spacingPercent13 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing12.Append(spacingPercent13);
            A.DefaultRunProperties defaultRunProperties66 = new A.DefaultRunProperties();

            level1ParagraphProperties32.Append(lineSpacing12);
            level1ParagraphProperties32.Append(defaultRunProperties66);

            A.Level2ParagraphProperties level2ParagraphProperties5 = new A.Level2ParagraphProperties();

            A.LineSpacing lineSpacing13 = new A.LineSpacing();
            A.SpacingPercent spacingPercent14 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing13.Append(spacingPercent14);
            A.DefaultRunProperties defaultRunProperties67 = new A.DefaultRunProperties();

            level2ParagraphProperties5.Append(lineSpacing13);
            level2ParagraphProperties5.Append(defaultRunProperties67);

            A.Level3ParagraphProperties level3ParagraphProperties5 = new A.Level3ParagraphProperties();

            A.LineSpacing lineSpacing14 = new A.LineSpacing();
            A.SpacingPercent spacingPercent15 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing14.Append(spacingPercent15);
            A.DefaultRunProperties defaultRunProperties68 = new A.DefaultRunProperties();

            level3ParagraphProperties5.Append(lineSpacing14);
            level3ParagraphProperties5.Append(defaultRunProperties68);

            listStyle54.Append(level1ParagraphProperties32);
            listStyle54.Append(level2ParagraphProperties5);
            listStyle54.Append(level3ParagraphProperties5);

            A.Paragraph paragraph60 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties28 = new A.ParagraphProperties() { Level = 0 };

            A.Run run58 = new A.Run();
            A.RunProperties runProperties61 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text60 = new A.Text();
            text60.Text = "Click to edit Master text styles";

            run58.Append(runProperties61);
            run58.Append(text60);

            paragraph60.Append(paragraphProperties28);
            paragraph60.Append(run58);

            A.Paragraph paragraph61 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties29 = new A.ParagraphProperties() { Level = 1 };

            A.Run run59 = new A.Run();
            A.RunProperties runProperties62 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text61 = new A.Text();
            text61.Text = "Second level";

            run59.Append(runProperties62);
            run59.Append(text61);

            paragraph61.Append(paragraphProperties29);
            paragraph61.Append(run59);

            A.Paragraph paragraph62 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties30 = new A.ParagraphProperties() { Level = 2 };

            A.Run run60 = new A.Run();
            A.RunProperties runProperties63 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text62 = new A.Text();
            text62.Text = "Third level";

            run60.Append(runProperties63);
            run60.Append(text62);

            paragraph62.Append(paragraphProperties30);
            paragraph62.Append(run60);

            textBody54.Append(bodyProperties54);
            textBody54.Append(listStyle54);
            textBody54.Append(paragraph60);
            textBody54.Append(paragraph61);
            textBody54.Append(paragraph62);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties65);
            shape59.Append(textBody54);

            shapeTree6.Append(nonVisualGroupShapeProperties8);
            shapeTree6.Append(groupShapeProperties8);
            shapeTree6.Append(shape58);
            shapeTree6.Append(shape59);

            CommonSlideDataExtensionList commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension4 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId4 = new P14.CreationId() { Val = (UInt32Value)530271944U };
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData6.Append(shapeTree6);
            commonSlideData6.Append(commonSlideDataExtensionList4);

            ColorMapOverride colorMapOverride4 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping4 = new A.MasterColorMapping();

            colorMapOverride4.Append(masterColorMapping4);

            slideLayout3.Append(commonSlideData6);
            slideLayout3.Append(colorMapOverride4);

            slideLayoutPart3.SlideLayout = slideLayout3;
        }

        // Generates content of slideLayoutPart4.
        private void GenerateSlideLayoutPart4Content(SlideLayoutPart slideLayoutPart4)
        {
            SlideLayout slideLayout4 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData7 = new CommonSlideData() { Name = "Headline 3 (Colour)" };

            ShapeTree shapeTree7 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties9 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties74 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties9 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties74 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties9.Append(nonVisualDrawingProperties74);
            nonVisualGroupShapeProperties9.Append(nonVisualGroupShapeDrawingProperties9);
            nonVisualGroupShapeProperties9.Append(applicationNonVisualDrawingProperties74);

            GroupShapeProperties groupShapeProperties9 = new GroupShapeProperties();

            A.TransformGroup transformGroup9 = new A.TransformGroup();
            A.Offset offset61 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents61 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset9 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents9 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup9.Append(offset61);
            transformGroup9.Append(extents61);
            transformGroup9.Append(childOffset9);
            transformGroup9.Append(childExtents9);

            groupShapeProperties9.Append(transformGroup9);

            Picture picture5 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties5 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties75 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties5 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks5 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties5.Append(pictureLocks5);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties75 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties5.Append(nonVisualDrawingProperties75);
            nonVisualPictureProperties5.Append(nonVisualPictureDrawingProperties5);
            nonVisualPictureProperties5.Append(applicationNonVisualDrawingProperties75);

            BlipFill blipFill5 = new BlipFill();

            A.Blip blip5 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList5 = new A.BlipExtensionList();

            A.BlipExtension blipExtension5 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi4 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi4.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension5.Append(useLocalDpi4);

            blipExtensionList5.Append(blipExtension5);

            blip5.Append(blipExtensionList5);

            A.Stretch stretch5 = new A.Stretch();
            A.FillRectangle fillRectangle4 = new A.FillRectangle();

            stretch5.Append(fillRectangle4);

            blipFill5.Append(blip5);
            blipFill5.Append(stretch5);

            ShapeProperties shapeProperties66 = new ShapeProperties();

            A.Transform2D transform2D53 = new A.Transform2D();
            A.Offset offset62 = new A.Offset() { X = -16772L, Y = 0L };
            A.Extents extents62 = new A.Extents() { Cx = 12208772L, Cy = 6858000L };

            transform2D53.Append(offset62);
            transform2D53.Append(extents62);

            A.PresetGeometry presetGeometry30 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList30 = new A.AdjustValueList();

            presetGeometry30.Append(adjustValueList30);

            shapeProperties66.Append(transform2D53);
            shapeProperties66.Append(presetGeometry30);

            picture5.Append(nonVisualPictureProperties5);
            picture5.Append(blipFill5);
            picture5.Append(shapeProperties66);

            Shape shape60 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties60 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties76 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Placeholder 12" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties60 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks50 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties60.Append(shapeLocks50);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties76 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape50 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties76.Append(placeholderShape50);

            nonVisualShapeProperties60.Append(nonVisualDrawingProperties76);
            nonVisualShapeProperties60.Append(nonVisualShapeDrawingProperties60);
            nonVisualShapeProperties60.Append(applicationNonVisualDrawingProperties76);

            ShapeProperties shapeProperties67 = new ShapeProperties();

            A.Transform2D transform2D54 = new A.Transform2D();
            A.Offset offset63 = new A.Offset() { X = -17463L, Y = 0L };
            A.Extents extents63 = new A.Extents() { Cx = 12209463L, Cy = 6858000L };

            transform2D54.Append(offset63);
            transform2D54.Append(extents63);

            A.SolidFill solidFill67 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex23 = new A.RgbColorModelHex() { Val = "86D5AC" };
            A.Alpha alpha4 = new A.Alpha() { Val = 90000 };

            rgbColorModelHex23.Append(alpha4);

            solidFill67.Append(rgbColorModelHex23);

            shapeProperties67.Append(transform2D54);
            shapeProperties67.Append(solidFill67);

            TextBody textBody55 = new TextBody();

            A.BodyProperties bodyProperties55 = new A.BodyProperties();
            A.NormalAutoFit normalAutoFit4 = new A.NormalAutoFit();

            bodyProperties55.Append(normalAutoFit4);

            A.ListStyle listStyle55 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties33 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet22 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties69 = new A.DefaultRunProperties() { FontSize = 1400 };

            A.SolidFill solidFill68 = new A.SolidFill();

            A.SchemeColor schemeColor99 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.Alpha alpha5 = new A.Alpha() { Val = 0 };

            schemeColor99.Append(alpha5);

            solidFill68.Append(schemeColor99);

            defaultRunProperties69.Append(solidFill68);

            level1ParagraphProperties33.Append(noBullet22);
            level1ParagraphProperties33.Append(defaultRunProperties69);

            listStyle55.Append(level1ParagraphProperties33);

            A.Paragraph paragraph63 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties31 = new A.ParagraphProperties() { Level = 0 };

            A.Run run61 = new A.Run();
            A.RunProperties runProperties64 = new A.RunProperties() { Language = "en-US" };
            A.Text text63 = new A.Text();
            text63.Text = "Click to edit Master text styles";

            run61.Append(runProperties64);
            run61.Append(text63);

            paragraph63.Append(paragraphProperties31);
            paragraph63.Append(run61);

            textBody55.Append(bodyProperties55);
            textBody55.Append(listStyle55);
            textBody55.Append(paragraph63);

            shape60.Append(nonVisualShapeProperties60);
            shape60.Append(shapeProperties67);
            shape60.Append(textBody55);

            Shape shape61 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties61 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties77 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties61 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks51 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties61.Append(shapeLocks51);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties77 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape51 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties77.Append(placeholderShape51);

            nonVisualShapeProperties61.Append(nonVisualDrawingProperties77);
            nonVisualShapeProperties61.Append(nonVisualShapeDrawingProperties61);
            nonVisualShapeProperties61.Append(applicationNonVisualDrawingProperties77);

            ShapeProperties shapeProperties68 = new ShapeProperties();

            A.Transform2D transform2D55 = new A.Transform2D();
            A.Offset offset64 = new A.Offset() { X = 695325L, Y = 2934585L };
            A.Extents extents64 = new A.Extents() { Cx = 10801350L, Cy = 952193L };

            transform2D55.Append(offset64);
            transform2D55.Append(extents64);

            shapeProperties68.Append(transform2D55);

            TextBody textBody56 = new TextBody();
            A.BodyProperties bodyProperties56 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle56 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties34 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties70 = new A.DefaultRunProperties() { FontSize = 4400 };

            A.SolidFill solidFill69 = new A.SolidFill();
            A.SchemeColor schemeColor100 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill69.Append(schemeColor100);

            defaultRunProperties70.Append(solidFill69);

            level1ParagraphProperties34.Append(defaultRunProperties70);

            listStyle56.Append(level1ParagraphProperties34);

            A.Paragraph paragraph64 = new A.Paragraph();

            A.Run run62 = new A.Run();
            A.RunProperties runProperties65 = new A.RunProperties() { Language = "en-US" };
            A.Text text64 = new A.Text();
            text64.Text = "Click to edit Master title style";

            run62.Append(runProperties65);
            run62.Append(text64);

            paragraph64.Append(run62);

            textBody56.Append(bodyProperties56);
            textBody56.Append(listStyle56);
            textBody56.Append(paragraph64);

            shape61.Append(nonVisualShapeProperties61);
            shape61.Append(shapeProperties68);
            shape61.Append(textBody56);

            shapeTree7.Append(nonVisualGroupShapeProperties9);
            shapeTree7.Append(groupShapeProperties9);
            shapeTree7.Append(picture5);
            shapeTree7.Append(shape60);
            shapeTree7.Append(shape61);

            commonSlideData7.Append(shapeTree7);

            ColorMapOverride colorMapOverride5 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping5 = new A.MasterColorMapping();

            colorMapOverride5.Append(masterColorMapping5);

            slideLayout4.Append(commonSlideData7);
            slideLayout4.Append(colorMapOverride5);

            slideLayoutPart4.SlideLayout = slideLayout4;
        }

        // Generates content of slideLayoutPart5.
        private void GenerateSlideLayoutPart5Content(SlideLayoutPart slideLayoutPart5)
        {
            SlideLayout slideLayout5 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData8 = new CommonSlideData() { Name = "Cover 2" };

            ShapeTree shapeTree8 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties78 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties78 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties78);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties78);

            GroupShapeProperties groupShapeProperties10 = new GroupShapeProperties();

            A.TransformGroup transformGroup10 = new A.TransformGroup();
            A.Offset offset65 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents65 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset10 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents10 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup10.Append(offset65);
            transformGroup10.Append(extents65);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            Picture picture6 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties6 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties79 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture 2" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties6 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks6 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties6.Append(pictureLocks6);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties79 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties6.Append(nonVisualDrawingProperties79);
            nonVisualPictureProperties6.Append(nonVisualPictureDrawingProperties6);
            nonVisualPictureProperties6.Append(applicationNonVisualDrawingProperties79);

            BlipFill blipFill6 = new BlipFill();

            A.Blip blip6 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList6 = new A.BlipExtensionList();

            A.BlipExtension blipExtension6 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi5 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi5.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension6.Append(useLocalDpi5);

            blipExtensionList6.Append(blipExtension6);

            blip6.Append(blipExtensionList6);

            A.Stretch stretch6 = new A.Stretch();
            A.FillRectangle fillRectangle5 = new A.FillRectangle();

            stretch6.Append(fillRectangle5);

            blipFill6.Append(blip6);
            blipFill6.Append(stretch6);

            ShapeProperties shapeProperties69 = new ShapeProperties();

            A.Transform2D transform2D56 = new A.Transform2D();
            A.Offset offset66 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents66 = new A.Extents() { Cx = 12195050L, Cy = 6858000L };

            transform2D56.Append(offset66);
            transform2D56.Append(extents66);

            A.PresetGeometry presetGeometry31 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList31 = new A.AdjustValueList();

            presetGeometry31.Append(adjustValueList31);

            shapeProperties69.Append(transform2D56);
            shapeProperties69.Append(presetGeometry31);

            picture6.Append(nonVisualPictureProperties6);
            picture6.Append(blipFill6);
            picture6.Append(shapeProperties69);

            Picture picture7 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties7 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties80 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties7 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks7 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties7.Append(pictureLocks7);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties80 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties7.Append(nonVisualDrawingProperties80);
            nonVisualPictureProperties7.Append(nonVisualPictureDrawingProperties7);
            nonVisualPictureProperties7.Append(applicationNonVisualDrawingProperties80);

            BlipFill blipFill7 = new BlipFill();
            A.Blip blip7 = new A.Blip() { Embed = "rId2" };

            A.Stretch stretch7 = new A.Stretch();
            A.FillRectangle fillRectangle6 = new A.FillRectangle();

            stretch7.Append(fillRectangle6);

            blipFill7.Append(blip7);
            blipFill7.Append(stretch7);

            ShapeProperties shapeProperties70 = new ShapeProperties();

            A.Transform2D transform2D57 = new A.Transform2D();
            A.Offset offset67 = new A.Offset() { X = 4151011L, Y = 2592729L };
            A.Extents extents67 = new A.Extents() { Cx = 4229685L, Cy = 1010534L };

            transform2D57.Append(offset67);
            transform2D57.Append(extents67);

            A.PresetGeometry presetGeometry32 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList32 = new A.AdjustValueList();

            presetGeometry32.Append(adjustValueList32);

            shapeProperties70.Append(transform2D57);
            shapeProperties70.Append(presetGeometry32);

            picture7.Append(nonVisualPictureProperties7);
            picture7.Append(blipFill7);
            picture7.Append(shapeProperties70);

            shapeTree8.Append(nonVisualGroupShapeProperties10);
            shapeTree8.Append(groupShapeProperties10);
            shapeTree8.Append(picture6);
            shapeTree8.Append(picture7);
            CommonSlideDataExtensionList commonSlideDataExtensionList5 = new CommonSlideDataExtensionList();

            commonSlideData8.Append(shapeTree8);
            commonSlideData8.Append(commonSlideDataExtensionList5);

            ColorMapOverride colorMapOverride6 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping6 = new A.MasterColorMapping();

            colorMapOverride6.Append(masterColorMapping6);

            slideLayout5.Append(commonSlideData8);
            slideLayout5.Append(colorMapOverride6);

            slideLayoutPart5.SlideLayout = slideLayout5;
        }

        // Generates content of slideLayoutPart6.
        private void GenerateSlideLayoutPart6Content(SlideLayoutPart slideLayoutPart6)
        {
            SlideLayout slideLayout6 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData9 = new CommonSlideData() { Name = "Cover" };

            ShapeTree shapeTree9 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties11 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties81 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties11 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties81 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties11.Append(nonVisualDrawingProperties81);
            nonVisualGroupShapeProperties11.Append(nonVisualGroupShapeDrawingProperties11);
            nonVisualGroupShapeProperties11.Append(applicationNonVisualDrawingProperties81);

            GroupShapeProperties groupShapeProperties11 = new GroupShapeProperties();

            A.TransformGroup transformGroup11 = new A.TransformGroup();
            A.Offset offset68 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents68 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset11 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents11 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup11.Append(offset68);
            transformGroup11.Append(extents68);
            transformGroup11.Append(childOffset11);
            transformGroup11.Append(childExtents11);

            groupShapeProperties11.Append(transformGroup11);

            Picture picture8 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties8 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties82 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture 2" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties8 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks8 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties8.Append(pictureLocks8);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties82 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties8.Append(nonVisualDrawingProperties82);
            nonVisualPictureProperties8.Append(nonVisualPictureDrawingProperties8);
            nonVisualPictureProperties8.Append(applicationNonVisualDrawingProperties82);

            BlipFill blipFill8 = new BlipFill();

            A.Blip blip8 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList7 = new A.BlipExtensionList();

            A.BlipExtension blipExtension7 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi6 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi6.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension7.Append(useLocalDpi6);

            blipExtensionList7.Append(blipExtension7);

            blip8.Append(blipExtensionList7);

            A.Stretch stretch8 = new A.Stretch();
            A.FillRectangle fillRectangle7 = new A.FillRectangle();

            stretch8.Append(fillRectangle7);

            blipFill8.Append(blip8);
            blipFill8.Append(stretch8);

            ShapeProperties shapeProperties71 = new ShapeProperties();

            A.Transform2D transform2D58 = new A.Transform2D();
            A.Offset offset69 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents69 = new A.Extents() { Cx = 12195050L, Cy = 6858000L };

            transform2D58.Append(offset69);
            transform2D58.Append(extents69);

            A.PresetGeometry presetGeometry33 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList33 = new A.AdjustValueList();

            presetGeometry33.Append(adjustValueList33);

            shapeProperties71.Append(transform2D58);
            shapeProperties71.Append(presetGeometry33);

            picture8.Append(nonVisualPictureProperties8);
            picture8.Append(blipFill8);
            picture8.Append(shapeProperties71);

            Picture picture9 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties9 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties83 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties9 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks9 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties9.Append(pictureLocks9);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties83 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties9.Append(nonVisualDrawingProperties83);
            nonVisualPictureProperties9.Append(nonVisualPictureDrawingProperties9);
            nonVisualPictureProperties9.Append(applicationNonVisualDrawingProperties83);

            BlipFill blipFill9 = new BlipFill();
            A.Blip blip9 = new A.Blip() { Embed = "rId2" };

            A.Stretch stretch9 = new A.Stretch();
            A.FillRectangle fillRectangle8 = new A.FillRectangle();

            stretch9.Append(fillRectangle8);

            blipFill9.Append(blip9);
            blipFill9.Append(stretch9);

            ShapeProperties shapeProperties72 = new ShapeProperties();

            A.Transform2D transform2D59 = new A.Transform2D();
            A.Offset offset70 = new A.Offset() { X = 4151011L, Y = 2592729L };
            A.Extents extents70 = new A.Extents() { Cx = 4229685L, Cy = 1010534L };

            transform2D59.Append(offset70);
            transform2D59.Append(extents70);

            A.PresetGeometry presetGeometry34 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList34 = new A.AdjustValueList();

            presetGeometry34.Append(adjustValueList34);

            shapeProperties72.Append(transform2D59);
            shapeProperties72.Append(presetGeometry34);

            picture9.Append(nonVisualPictureProperties9);
            picture9.Append(blipFill9);
            picture9.Append(shapeProperties72);

            GroupShape groupShape3 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties84 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Group 4" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties84 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties84);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties84);

            GroupShapeProperties groupShapeProperties12 = new GroupShapeProperties();

            A.TransformGroup transformGroup12 = new A.TransformGroup();
            A.Offset offset71 = new A.Offset() { X = 10356783L, Y = 2945331L };
            A.Extents extents71 = new A.Extents() { Cx = 1139892L, Cy = 1139892L };
            A.ChildOffset childOffset12 = new A.ChildOffset() { X = 10356783L, Y = 2945331L };
            A.ChildExtents childExtents12 = new A.ChildExtents() { Cx = 1139892L, Cy = 1139892L };

            transformGroup12.Append(offset71);
            transformGroup12.Append(extents71);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            Shape shape62 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties62 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties85 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Oval 5" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties62 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties85 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties62.Append(nonVisualDrawingProperties85);
            nonVisualShapeProperties62.Append(nonVisualShapeDrawingProperties62);
            nonVisualShapeProperties62.Append(applicationNonVisualDrawingProperties85);

            ShapeProperties shapeProperties73 = new ShapeProperties();

            A.Transform2D transform2D60 = new A.Transform2D();
            A.Offset offset72 = new A.Offset() { X = 10356783L, Y = 2945331L };
            A.Extents extents72 = new A.Extents() { Cx = 1139892L, Cy = 1139892L };

            transform2D60.Append(offset72);
            transform2D60.Append(extents72);

            A.PresetGeometry presetGeometry35 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList35 = new A.AdjustValueList();

            presetGeometry35.Append(adjustValueList35);

            A.SolidFill solidFill70 = new A.SolidFill();
            A.SchemeColor schemeColor101 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill70.Append(schemeColor101);

            A.Outline outline12 = new A.Outline();
            A.NoFill noFill15 = new A.NoFill();

            outline12.Append(noFill15);

            shapeProperties73.Append(transform2D60);
            shapeProperties73.Append(presetGeometry35);
            shapeProperties73.Append(solidFill70);
            shapeProperties73.Append(outline12);

            ShapeStyle shapeStyle7 = new ShapeStyle();

            A.LineReference lineReference7 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor102 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade10 = new A.Shade() { Val = 50000 };

            schemeColor102.Append(shade10);

            lineReference7.Append(schemeColor102);

            A.FillReference fillReference7 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor103 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference7.Append(schemeColor103);

            A.EffectReference effectReference7 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor104 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference7.Append(schemeColor104);

            A.FontReference fontReference7 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor105 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference7.Append(schemeColor105);

            shapeStyle7.Append(lineReference7);
            shapeStyle7.Append(fillReference7);
            shapeStyle7.Append(effectReference7);
            shapeStyle7.Append(fontReference7);

            TextBody textBody57 = new TextBody();
            A.BodyProperties bodyProperties57 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle57 = new A.ListStyle();

            A.Paragraph paragraph65 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties24 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph65.Append(paragraphProperties32);
            paragraph65.Append(endParagraphRunProperties24);

            textBody57.Append(bodyProperties57);
            textBody57.Append(listStyle57);
            textBody57.Append(paragraph65);

            shape62.Append(nonVisualShapeProperties62);
            shape62.Append(shapeProperties73);
            shape62.Append(shapeStyle7);
            shape62.Append(textBody57);

            GroupShape groupShape4 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties13 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties86 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Group 6" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties13 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties86 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties13.Append(nonVisualDrawingProperties86);
            nonVisualGroupShapeProperties13.Append(nonVisualGroupShapeDrawingProperties13);
            nonVisualGroupShapeProperties13.Append(applicationNonVisualDrawingProperties86);

            GroupShapeProperties groupShapeProperties13 = new GroupShapeProperties();

            A.TransformGroup transformGroup13 = new A.TransformGroup();
            A.Offset offset73 = new A.Offset() { X = 10804795L, Y = 3324431L };
            A.Extents extents73 = new A.Extents() { Cx = 243867L, Cy = 381692L };
            A.ChildOffset childOffset13 = new A.ChildOffset() { X = 6996222L, Y = 4465674L };
            A.ChildExtents childExtents13 = new A.ChildExtents() { Cx = 122276L, Cy = 191386L };

            transformGroup13.Append(offset73);
            transformGroup13.Append(extents73);
            transformGroup13.Append(childOffset13);
            transformGroup13.Append(childExtents13);

            groupShapeProperties13.Append(transformGroup13);

            ConnectionShape connectionShape3 = new ConnectionShape();

            NonVisualConnectionShapeProperties nonVisualConnectionShapeProperties3 = new NonVisualConnectionShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties87 = new NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Straight Connector 7" };
            NonVisualConnectorShapeDrawingProperties nonVisualConnectorShapeDrawingProperties3 = new NonVisualConnectorShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties87 = new ApplicationNonVisualDrawingProperties();

            nonVisualConnectionShapeProperties3.Append(nonVisualDrawingProperties87);
            nonVisualConnectionShapeProperties3.Append(nonVisualConnectorShapeDrawingProperties3);
            nonVisualConnectionShapeProperties3.Append(applicationNonVisualDrawingProperties87);

            ShapeProperties shapeProperties74 = new ShapeProperties();

            A.Transform2D transform2D61 = new A.Transform2D();
            A.Offset offset74 = new A.Offset() { X = 6996223L, Y = 4465674L };
            A.Extents extents74 = new A.Extents() { Cx = 122275L, Cy = 95693L };

            transform2D61.Append(offset74);
            transform2D61.Append(extents74);

            A.PresetGeometry presetGeometry36 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Line };
            A.AdjustValueList adjustValueList36 = new A.AdjustValueList();

            presetGeometry36.Append(adjustValueList36);

            A.Outline outline13 = new A.Outline() { Width = 44450, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill71 = new A.SolidFill();
            A.SchemeColor schemeColor106 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill71.Append(schemeColor106);
            A.Round round3 = new A.Round();

            outline13.Append(solidFill71);
            outline13.Append(round3);

            shapeProperties74.Append(transform2D61);
            shapeProperties74.Append(presetGeometry36);
            shapeProperties74.Append(outline13);

            ShapeStyle shapeStyle8 = new ShapeStyle();

            A.LineReference lineReference8 = new A.LineReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor107 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference8.Append(schemeColor107);

            A.FillReference fillReference8 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor108 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference8.Append(schemeColor108);

            A.EffectReference effectReference8 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor109 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference8.Append(schemeColor109);

            A.FontReference fontReference8 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor110 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference8.Append(schemeColor110);

            shapeStyle8.Append(lineReference8);
            shapeStyle8.Append(fillReference8);
            shapeStyle8.Append(effectReference8);
            shapeStyle8.Append(fontReference8);

            connectionShape3.Append(nonVisualConnectionShapeProperties3);
            connectionShape3.Append(shapeProperties74);
            connectionShape3.Append(shapeStyle8);

            ConnectionShape connectionShape4 = new ConnectionShape();

            NonVisualConnectionShapeProperties nonVisualConnectionShapeProperties4 = new NonVisualConnectionShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties88 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Straight Connector 8" };
            NonVisualConnectorShapeDrawingProperties nonVisualConnectorShapeDrawingProperties4 = new NonVisualConnectorShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties88 = new ApplicationNonVisualDrawingProperties();

            nonVisualConnectionShapeProperties4.Append(nonVisualDrawingProperties88);
            nonVisualConnectionShapeProperties4.Append(nonVisualConnectorShapeDrawingProperties4);
            nonVisualConnectionShapeProperties4.Append(applicationNonVisualDrawingProperties88);

            ShapeProperties shapeProperties75 = new ShapeProperties();

            A.Transform2D transform2D62 = new A.Transform2D() { HorizontalFlip = true };
            A.Offset offset75 = new A.Offset() { X = 6996222L, Y = 4561367L };
            A.Extents extents75 = new A.Extents() { Cx = 122276L, Cy = 95693L };

            transform2D62.Append(offset75);
            transform2D62.Append(extents75);

            A.PresetGeometry presetGeometry37 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Line };
            A.AdjustValueList adjustValueList37 = new A.AdjustValueList();

            presetGeometry37.Append(adjustValueList37);

            A.Outline outline14 = new A.Outline() { Width = 44450, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill72 = new A.SolidFill();
            A.SchemeColor schemeColor111 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill72.Append(schemeColor111);
            A.Round round4 = new A.Round();

            outline14.Append(solidFill72);
            outline14.Append(round4);

            shapeProperties75.Append(transform2D62);
            shapeProperties75.Append(presetGeometry37);
            shapeProperties75.Append(outline14);

            ShapeStyle shapeStyle9 = new ShapeStyle();

            A.LineReference lineReference9 = new A.LineReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor112 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference9.Append(schemeColor112);

            A.FillReference fillReference9 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor113 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference9.Append(schemeColor113);

            A.EffectReference effectReference9 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor114 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference9.Append(schemeColor114);

            A.FontReference fontReference9 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor115 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference9.Append(schemeColor115);

            shapeStyle9.Append(lineReference9);
            shapeStyle9.Append(fillReference9);
            shapeStyle9.Append(effectReference9);
            shapeStyle9.Append(fontReference9);

            connectionShape4.Append(nonVisualConnectionShapeProperties4);
            connectionShape4.Append(shapeProperties75);
            connectionShape4.Append(shapeStyle9);

            groupShape4.Append(nonVisualGroupShapeProperties13);
            groupShape4.Append(groupShapeProperties13);
            groupShape4.Append(connectionShape3);
            groupShape4.Append(connectionShape4);

            groupShape3.Append(nonVisualGroupShapeProperties12);
            groupShape3.Append(groupShapeProperties12);
            groupShape3.Append(shape62);
            groupShape3.Append(groupShape4);

            Shape shape63 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties63 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties89 = new NonVisualDrawingProperties() { Id = (UInt32Value)10U, Name = "Oval 9" };
            A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://hlinkshowjump?jump=nextslide" };

            nonVisualDrawingProperties89.Append(hyperlinkOnClick1);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties63 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties89 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualShapeProperties63.Append(nonVisualDrawingProperties89);
            nonVisualShapeProperties63.Append(nonVisualShapeDrawingProperties63);
            nonVisualShapeProperties63.Append(applicationNonVisualDrawingProperties89);

            ShapeProperties shapeProperties76 = new ShapeProperties();

            A.Transform2D transform2D63 = new A.Transform2D();
            A.Offset offset76 = new A.Offset() { X = 10356782L, Y = 2945331L };
            A.Extents extents76 = new A.Extents() { Cx = 1139892L, Cy = 1139892L };

            transform2D63.Append(offset76);
            transform2D63.Append(extents76);

            A.PresetGeometry presetGeometry38 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList38 = new A.AdjustValueList();

            presetGeometry38.Append(adjustValueList38);

            A.SolidFill solidFill73 = new A.SolidFill();

            A.SchemeColor schemeColor116 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Alpha alpha6 = new A.Alpha() { Val = 0 };

            schemeColor116.Append(alpha6);

            solidFill73.Append(schemeColor116);

            A.Outline outline15 = new A.Outline();
            A.NoFill noFill16 = new A.NoFill();

            outline15.Append(noFill16);

            shapeProperties76.Append(transform2D63);
            shapeProperties76.Append(presetGeometry38);
            shapeProperties76.Append(solidFill73);
            shapeProperties76.Append(outline15);

            ShapeStyle shapeStyle10 = new ShapeStyle();

            A.LineReference lineReference10 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor117 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade11 = new A.Shade() { Val = 50000 };

            schemeColor117.Append(shade11);

            lineReference10.Append(schemeColor117);

            A.FillReference fillReference10 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor118 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference10.Append(schemeColor118);

            A.EffectReference effectReference10 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor119 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference10.Append(schemeColor119);

            A.FontReference fontReference10 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor120 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference10.Append(schemeColor120);

            shapeStyle10.Append(lineReference10);
            shapeStyle10.Append(fillReference10);
            shapeStyle10.Append(effectReference10);
            shapeStyle10.Append(fontReference10);

            TextBody textBody58 = new TextBody();
            A.BodyProperties bodyProperties58 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle58 = new A.ListStyle();

            A.Paragraph paragraph66 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties33 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties25 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph66.Append(paragraphProperties33);
            paragraph66.Append(endParagraphRunProperties25);

            textBody58.Append(bodyProperties58);
            textBody58.Append(listStyle58);
            textBody58.Append(paragraph66);

            shape63.Append(nonVisualShapeProperties63);
            shape63.Append(shapeProperties76);
            shape63.Append(shapeStyle10);
            shape63.Append(textBody58);

            shapeTree9.Append(nonVisualGroupShapeProperties11);
            shapeTree9.Append(groupShapeProperties11);
            shapeTree9.Append(picture8);
            shapeTree9.Append(picture9);
            shapeTree9.Append(groupShape3);
            shapeTree9.Append(shape63);

            CommonSlideDataExtensionList commonSlideDataExtensionList6 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension5 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId5 = new P14.CreationId() { Val = (UInt32Value)1377184672U };
            creationId5.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension5.Append(creationId5);

            commonSlideDataExtensionList6.Append(commonSlideDataExtension5);

            commonSlideData9.Append(shapeTree9);
            commonSlideData9.Append(commonSlideDataExtensionList6);

            ColorMapOverride colorMapOverride7 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping7 = new A.MasterColorMapping();

            colorMapOverride7.Append(masterColorMapping7);

            slideLayout6.Append(commonSlideData9);
            slideLayout6.Append(colorMapOverride7);

            slideLayoutPart6.SlideLayout = slideLayout6;
        }

        // Generates content of slideLayoutPart7.
        private void GenerateSlideLayoutPart7Content(SlideLayoutPart slideLayoutPart7)
        {
            SlideLayout slideLayout7 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout7.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout7.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData10 = new CommonSlideData() { Name = "Headline 2" };

            ShapeTree shapeTree10 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties14 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties90 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties14 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties90 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties14.Append(nonVisualDrawingProperties90);
            nonVisualGroupShapeProperties14.Append(nonVisualGroupShapeDrawingProperties14);
            nonVisualGroupShapeProperties14.Append(applicationNonVisualDrawingProperties90);

            GroupShapeProperties groupShapeProperties14 = new GroupShapeProperties();

            A.TransformGroup transformGroup14 = new A.TransformGroup();
            A.Offset offset77 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents77 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset14 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents14 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup14.Append(offset77);
            transformGroup14.Append(extents77);
            transformGroup14.Append(childOffset14);
            transformGroup14.Append(childExtents14);

            groupShapeProperties14.Append(transformGroup14);

            Picture picture10 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties10 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties91 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties10 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks10 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties10.Append(pictureLocks10);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties91 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties10.Append(nonVisualDrawingProperties91);
            nonVisualPictureProperties10.Append(nonVisualPictureDrawingProperties10);
            nonVisualPictureProperties10.Append(applicationNonVisualDrawingProperties91);

            BlipFill blipFill10 = new BlipFill();

            A.Blip blip10 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList8 = new A.BlipExtensionList();

            A.BlipExtension blipExtension8 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi7 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi7.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension8.Append(useLocalDpi7);

            blipExtensionList8.Append(blipExtension8);

            blip10.Append(blipExtensionList8);

            A.Stretch stretch10 = new A.Stretch();
            A.FillRectangle fillRectangle9 = new A.FillRectangle();

            stretch10.Append(fillRectangle9);

            blipFill10.Append(blip10);
            blipFill10.Append(stretch10);

            ShapeProperties shapeProperties77 = new ShapeProperties();

            A.Transform2D transform2D64 = new A.Transform2D();
            A.Offset offset78 = new A.Offset() { X = -16772L, Y = 0L };
            A.Extents extents78 = new A.Extents() { Cx = 12208772L, Cy = 6858000L };

            transform2D64.Append(offset78);
            transform2D64.Append(extents78);

            A.PresetGeometry presetGeometry39 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList39 = new A.AdjustValueList();

            presetGeometry39.Append(adjustValueList39);

            shapeProperties77.Append(transform2D64);
            shapeProperties77.Append(presetGeometry39);

            picture10.Append(nonVisualPictureProperties10);
            picture10.Append(blipFill10);
            picture10.Append(shapeProperties77);

            Shape shape64 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties64 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties92 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties64 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks52 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties64.Append(shapeLocks52);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties92 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape52 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties92.Append(placeholderShape52);

            nonVisualShapeProperties64.Append(nonVisualDrawingProperties92);
            nonVisualShapeProperties64.Append(nonVisualShapeDrawingProperties64);
            nonVisualShapeProperties64.Append(applicationNonVisualDrawingProperties92);

            ShapeProperties shapeProperties78 = new ShapeProperties();

            A.Transform2D transform2D65 = new A.Transform2D();
            A.Offset offset79 = new A.Offset() { X = 695325L, Y = 2934585L };
            A.Extents extents79 = new A.Extents() { Cx = 10801350L, Cy = 952193L };

            transform2D65.Append(offset79);
            transform2D65.Append(extents79);

            shapeProperties78.Append(transform2D65);

            TextBody textBody59 = new TextBody();
            A.BodyProperties bodyProperties59 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle59 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties35 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties71 = new A.DefaultRunProperties() { FontSize = 4400 };

            A.SolidFill solidFill74 = new A.SolidFill();
            A.SchemeColor schemeColor121 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill74.Append(schemeColor121);

            defaultRunProperties71.Append(solidFill74);

            level1ParagraphProperties35.Append(defaultRunProperties71);

            listStyle59.Append(level1ParagraphProperties35);

            A.Paragraph paragraph67 = new A.Paragraph();

            A.Run run63 = new A.Run();
            A.RunProperties runProperties66 = new A.RunProperties() { Language = "en-US" };
            A.Text text65 = new A.Text();
            text65.Text = "Click to edit Master title style";

            run63.Append(runProperties66);
            run63.Append(text65);

            paragraph67.Append(run63);

            textBody59.Append(bodyProperties59);
            textBody59.Append(listStyle59);
            textBody59.Append(paragraph67);

            shape64.Append(nonVisualShapeProperties64);
            shape64.Append(shapeProperties78);
            shape64.Append(textBody59);

            Picture picture11 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties11 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties93 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Picture 5" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties11 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks11 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties11.Append(pictureLocks11);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties93 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties11.Append(nonVisualDrawingProperties93);
            nonVisualPictureProperties11.Append(nonVisualPictureDrawingProperties11);
            nonVisualPictureProperties11.Append(applicationNonVisualDrawingProperties93);

            BlipFill blipFill11 = new BlipFill();

            A.Blip blip11 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList9 = new A.BlipExtensionList();

            A.BlipExtension blipExtension9 = new A.BlipExtension() { Uri = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}" };

            A14.ImageProperties imageProperties2 = new A14.ImageProperties();
            imageProperties2.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A14.ImageLayer imageLayer2 = new A14.ImageLayer() { Embed = "rId3" };

            A14.ImageEffect imageEffect2 = new A14.ImageEffect();
            A14.BrightnessContrast brightnessContrast2 = new A14.BrightnessContrast() { Bright = 100000 };

            imageEffect2.Append(brightnessContrast2);

            imageLayer2.Append(imageEffect2);

            imageProperties2.Append(imageLayer2);

            blipExtension9.Append(imageProperties2);

            A.BlipExtension blipExtension10 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi8 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi8.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension10.Append(useLocalDpi8);

            blipExtensionList9.Append(blipExtension9);
            blipExtensionList9.Append(blipExtension10);

            blip11.Append(blipExtensionList9);

            A.Stretch stretch11 = new A.Stretch();
            A.FillRectangle fillRectangle10 = new A.FillRectangle();

            stretch11.Append(fillRectangle10);

            blipFill11.Append(blip11);
            blipFill11.Append(stretch11);

            ShapeProperties shapeProperties79 = new ShapeProperties();

            A.Transform2D transform2D66 = new A.Transform2D();
            A.Offset offset80 = new A.Offset() { X = 10919356L, Y = 6465900L };
            A.Extents extents80 = new A.Extents() { Cx = 1095427L, Cy = 260968L };

            transform2D66.Append(offset80);
            transform2D66.Append(extents80);

            A.PresetGeometry presetGeometry40 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList40 = new A.AdjustValueList();

            presetGeometry40.Append(adjustValueList40);

            shapeProperties79.Append(transform2D66);
            shapeProperties79.Append(presetGeometry40);

            picture11.Append(nonVisualPictureProperties11);
            picture11.Append(blipFill11);
            picture11.Append(shapeProperties79);

            shapeTree10.Append(nonVisualGroupShapeProperties14);
            shapeTree10.Append(groupShapeProperties14);
            shapeTree10.Append(picture10);
            shapeTree10.Append(shape64);
            shapeTree10.Append(picture11);

            commonSlideData10.Append(shapeTree10);

            ColorMapOverride colorMapOverride8 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping8 = new A.MasterColorMapping();

            colorMapOverride8.Append(masterColorMapping8);

            slideLayout7.Append(commonSlideData10);
            slideLayout7.Append(colorMapOverride8);

            slideLayoutPart7.SlideLayout = slideLayout7;
        }

        // Generates content of slideLayoutPart8.
        private void GenerateSlideLayoutPart8Content(SlideLayoutPart slideLayoutPart8)
        {
            SlideLayout slideLayout8 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData11 = new CommonSlideData() { Name = "Headline 1" };

            ShapeTree shapeTree11 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties15 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties94 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties15 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties94 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties15.Append(nonVisualDrawingProperties94);
            nonVisualGroupShapeProperties15.Append(nonVisualGroupShapeDrawingProperties15);
            nonVisualGroupShapeProperties15.Append(applicationNonVisualDrawingProperties94);

            GroupShapeProperties groupShapeProperties15 = new GroupShapeProperties();

            A.TransformGroup transformGroup15 = new A.TransformGroup();
            A.Offset offset81 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents81 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset15 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents15 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup15.Append(offset81);
            transformGroup15.Append(extents81);
            transformGroup15.Append(childOffset15);
            transformGroup15.Append(childExtents15);

            groupShapeProperties15.Append(transformGroup15);

            Picture picture12 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties12 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties95 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Picture 2", Description = "https://calendar.jvbrown.edu/assets/Events/Photos/bubbles.png" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties12 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks12 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties12.Append(pictureLocks12);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties95 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties12.Append(nonVisualDrawingProperties95);
            nonVisualPictureProperties12.Append(nonVisualPictureDrawingProperties12);
            nonVisualPictureProperties12.Append(applicationNonVisualDrawingProperties95);

            BlipFill blipFill12 = new BlipFill() { RotateWithShape = true };

            A.Blip blip12 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList10 = new A.BlipExtensionList();

            A.BlipExtension blipExtension11 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi9 = new A14.UseLocalDpi();
            useLocalDpi9.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension11.Append(useLocalDpi9);

            blipExtensionList10.Append(blipExtension11);

            blip12.Append(blipExtensionList10);
            A.SourceRectangle sourceRectangle2 = new A.SourceRectangle() { Top = 12297, Bottom = 12530 };
            A.Stretch stretch12 = new A.Stretch();

            blipFill12.Append(blip12);
            blipFill12.Append(sourceRectangle2);
            blipFill12.Append(stretch12);

            ShapeProperties shapeProperties80 = new ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D67 = new A.Transform2D();
            A.Offset offset82 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents82 = new A.Extents() { Cx = 12191999L, Cy = 6858000L };

            transform2D67.Append(offset82);
            transform2D67.Append(extents82);

            A.PresetGeometry presetGeometry41 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList41 = new A.AdjustValueList();

            presetGeometry41.Append(adjustValueList41);
            A.NoFill noFill17 = new A.NoFill();

            A.Outline outline16 = new A.Outline();
            A.NoFill noFill18 = new A.NoFill();

            outline16.Append(noFill18);

            A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

            A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
            hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill75 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex24 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill75.Append(rgbColorModelHex24);

            hiddenFillProperties1.Append(solidFill75);

            shapePropertiesExtension1.Append(hiddenFillProperties1);

            shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

            shapeProperties80.Append(transform2D67);
            shapeProperties80.Append(presetGeometry41);
            shapeProperties80.Append(noFill17);
            shapeProperties80.Append(outline16);
            shapeProperties80.Append(shapePropertiesExtensionList1);

            picture12.Append(nonVisualPictureProperties12);
            picture12.Append(blipFill12);
            picture12.Append(shapeProperties80);

            Shape shape65 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties65 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties96 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties65 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks53 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties65.Append(shapeLocks53);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties96 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape53 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties96.Append(placeholderShape53);

            nonVisualShapeProperties65.Append(nonVisualDrawingProperties96);
            nonVisualShapeProperties65.Append(nonVisualShapeDrawingProperties65);
            nonVisualShapeProperties65.Append(applicationNonVisualDrawingProperties96);

            ShapeProperties shapeProperties81 = new ShapeProperties();

            A.Transform2D transform2D68 = new A.Transform2D();
            A.Offset offset83 = new A.Offset() { X = 695325L, Y = 2934585L };
            A.Extents extents83 = new A.Extents() { Cx = 10801350L, Cy = 952193L };

            transform2D68.Append(offset83);
            transform2D68.Append(extents83);

            shapeProperties81.Append(transform2D68);

            TextBody textBody60 = new TextBody();
            A.BodyProperties bodyProperties60 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle60 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties36 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties72 = new A.DefaultRunProperties() { FontSize = 4400 };

            A.SolidFill solidFill76 = new A.SolidFill();
            A.SchemeColor schemeColor122 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill76.Append(schemeColor122);

            defaultRunProperties72.Append(solidFill76);

            level1ParagraphProperties36.Append(defaultRunProperties72);

            listStyle60.Append(level1ParagraphProperties36);

            A.Paragraph paragraph68 = new A.Paragraph();

            A.Run run64 = new A.Run();
            A.RunProperties runProperties67 = new A.RunProperties() { Language = "en-US" };
            A.Text text66 = new A.Text();
            text66.Text = "Click to edit Master title style";

            run64.Append(runProperties67);
            run64.Append(text66);

            paragraph68.Append(run64);

            textBody60.Append(bodyProperties60);
            textBody60.Append(listStyle60);
            textBody60.Append(paragraph68);

            shape65.Append(nonVisualShapeProperties65);
            shape65.Append(shapeProperties81);
            shape65.Append(textBody60);

            Picture picture13 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties13 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties97 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Picture 5" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties13 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks13 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties13.Append(pictureLocks13);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties97 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties13.Append(nonVisualDrawingProperties97);
            nonVisualPictureProperties13.Append(nonVisualPictureDrawingProperties13);
            nonVisualPictureProperties13.Append(applicationNonVisualDrawingProperties97);

            BlipFill blipFill13 = new BlipFill();

            A.Blip blip13 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList11 = new A.BlipExtensionList();

            A.BlipExtension blipExtension12 = new A.BlipExtension() { Uri = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}" };

            A14.ImageProperties imageProperties3 = new A14.ImageProperties();
            imageProperties3.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A14.ImageLayer imageLayer3 = new A14.ImageLayer() { Embed = "rId3" };

            A14.ImageEffect imageEffect3 = new A14.ImageEffect();
            A14.BrightnessContrast brightnessContrast3 = new A14.BrightnessContrast() { Bright = 100000 };

            imageEffect3.Append(brightnessContrast3);

            imageLayer3.Append(imageEffect3);

            imageProperties3.Append(imageLayer3);

            blipExtension12.Append(imageProperties3);

            A.BlipExtension blipExtension13 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi10 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi10.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension13.Append(useLocalDpi10);

            blipExtensionList11.Append(blipExtension12);
            blipExtensionList11.Append(blipExtension13);

            blip13.Append(blipExtensionList11);

            A.Stretch stretch13 = new A.Stretch();
            A.FillRectangle fillRectangle11 = new A.FillRectangle();

            stretch13.Append(fillRectangle11);

            blipFill13.Append(blip13);
            blipFill13.Append(stretch13);

            ShapeProperties shapeProperties82 = new ShapeProperties();

            A.Transform2D transform2D69 = new A.Transform2D();
            A.Offset offset84 = new A.Offset() { X = 10919356L, Y = 6465900L };
            A.Extents extents84 = new A.Extents() { Cx = 1095427L, Cy = 260968L };

            transform2D69.Append(offset84);
            transform2D69.Append(extents84);

            A.PresetGeometry presetGeometry42 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList42 = new A.AdjustValueList();

            presetGeometry42.Append(adjustValueList42);

            shapeProperties82.Append(transform2D69);
            shapeProperties82.Append(presetGeometry42);

            picture13.Append(nonVisualPictureProperties13);
            picture13.Append(blipFill13);
            picture13.Append(shapeProperties82);

            shapeTree11.Append(nonVisualGroupShapeProperties15);
            shapeTree11.Append(groupShapeProperties15);
            shapeTree11.Append(picture12);
            shapeTree11.Append(shape65);
            shapeTree11.Append(picture13);

            commonSlideData11.Append(shapeTree11);

            ColorMapOverride colorMapOverride9 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout8.Append(commonSlideData11);
            slideLayout8.Append(colorMapOverride9);

            slideLayoutPart8.SlideLayout = slideLayout8;
        }

        // Generates content of themePart2.
        private void GenerateThemePart2Content(ThemePart themePart2)
        {
            A.Theme theme2 = new A.Theme() { Name = "Office Theme" };
            theme2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements2 = new A.ThemeElements();

            A.ColorScheme colorScheme2 = new A.ColorScheme() { Name = "DISCOVERAI" };

            A.Dark1Color dark1Color2 = new A.Dark1Color();
            A.RgbColorModelHex rgbColorModelHex25 = new A.RgbColorModelHex() { Val = "000000" };

            dark1Color2.Append(rgbColorModelHex25);

            A.Light1Color light1Color2 = new A.Light1Color();
            A.RgbColorModelHex rgbColorModelHex26 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            light1Color2.Append(rgbColorModelHex26);

            A.Dark2Color dark2Color2 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex27 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color2.Append(rgbColorModelHex27);

            A.Light2Color light2Color2 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex28 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color2.Append(rgbColorModelHex28);

            A.Accent1Color accent1Color2 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex29 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent1Color2.Append(rgbColorModelHex29);

            A.Accent2Color accent2Color2 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex30 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color2.Append(rgbColorModelHex30);

            A.Accent3Color accent3Color2 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex31 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color2.Append(rgbColorModelHex31);

            A.Accent4Color accent4Color2 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex32 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color2.Append(rgbColorModelHex32);

            A.Accent5Color accent5Color2 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex33 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent5Color2.Append(rgbColorModelHex33);

            A.Accent6Color accent6Color2 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex34 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color2.Append(rgbColorModelHex34);

            A.Hyperlink hyperlink2 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex35 = new A.RgbColorModelHex() { Val = "ABABAB" };

            hyperlink2.Append(rgbColorModelHex35);

            A.FollowedHyperlinkColor followedHyperlinkColor2 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex36 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor2.Append(rgbColorModelHex36);

            colorScheme2.Append(dark1Color2);
            colorScheme2.Append(light1Color2);
            colorScheme2.Append(dark2Color2);
            colorScheme2.Append(light2Color2);
            colorScheme2.Append(accent1Color2);
            colorScheme2.Append(accent2Color2);
            colorScheme2.Append(accent3Color2);
            colorScheme2.Append(accent4Color2);
            colorScheme2.Append(accent5Color2);
            colorScheme2.Append(accent6Color2);
            colorScheme2.Append(hyperlink2);
            colorScheme2.Append(followedHyperlinkColor2);

            A.FontScheme fontScheme2 = new A.FontScheme() { Name = "Arial" };

            A.MajorFont majorFont2 = new A.MajorFont();
            A.LatinFont latinFont62 = new A.LatinFont() { Typeface = "Arial", Panose = "020B0604020202020204" };
            A.EastAsianFont eastAsianFont62 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont62 = new A.ComplexScriptFont() { Typeface = "" };

            majorFont2.Append(latinFont62);
            majorFont2.Append(eastAsianFont62);
            majorFont2.Append(complexScriptFont62);

            A.MinorFont minorFont2 = new A.MinorFont();
            A.LatinFont latinFont63 = new A.LatinFont() { Typeface = "Arial", Panose = "020B0604020202020204" };
            A.EastAsianFont eastAsianFont63 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont63 = new A.ComplexScriptFont() { Typeface = "" };

            minorFont2.Append(latinFont63);
            minorFont2.Append(eastAsianFont63);
            minorFont2.Append(complexScriptFont63);

            fontScheme2.Append(majorFont2);
            fontScheme2.Append(minorFont2);

            A.FormatScheme formatScheme2 = new A.FormatScheme() { Name = "Office Theme" };

            A.FillStyleList fillStyleList2 = new A.FillStyleList();

            A.SolidFill solidFill77 = new A.SolidFill();
            A.SchemeColor schemeColor123 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill77.Append(schemeColor123);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor124 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint8 = new A.Tint() { Val = 67000 };

            schemeColor124.Append(luminanceModulation14);
            schemeColor124.Append(saturationModulation11);
            schemeColor124.Append(tint8);

            gradientStop10.Append(schemeColor124);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor125 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint9 = new A.Tint() { Val = 73000 };

            schemeColor125.Append(luminanceModulation15);
            schemeColor125.Append(saturationModulation12);
            schemeColor125.Append(tint9);

            gradientStop11.Append(schemeColor125);

            A.GradientStop gradientStop12 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor126 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation16 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation13 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint10 = new A.Tint() { Val = 81000 };

            schemeColor126.Append(luminanceModulation16);
            schemeColor126.Append(saturationModulation13);
            schemeColor126.Append(tint10);

            gradientStop12.Append(schemeColor126);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);
            gradientStopList4.Append(gradientStop12);
            A.LinearGradientFill linearGradientFill4 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(linearGradientFill4);

            A.GradientFill gradientFill5 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList5 = new A.GradientStopList();

            A.GradientStop gradientStop13 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor127 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation14 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation17 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint11 = new A.Tint() { Val = 94000 };

            schemeColor127.Append(saturationModulation14);
            schemeColor127.Append(luminanceModulation17);
            schemeColor127.Append(tint11);

            gradientStop13.Append(schemeColor127);

            A.GradientStop gradientStop14 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor128 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation15 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation18 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade12 = new A.Shade() { Val = 100000 };

            schemeColor128.Append(saturationModulation15);
            schemeColor128.Append(luminanceModulation18);
            schemeColor128.Append(shade12);

            gradientStop14.Append(schemeColor128);

            A.GradientStop gradientStop15 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor129 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation16 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade13 = new A.Shade() { Val = 78000 };

            schemeColor129.Append(luminanceModulation19);
            schemeColor129.Append(saturationModulation16);
            schemeColor129.Append(shade13);

            gradientStop15.Append(schemeColor129);

            gradientStopList5.Append(gradientStop13);
            gradientStopList5.Append(gradientStop14);
            gradientStopList5.Append(gradientStop15);
            A.LinearGradientFill linearGradientFill5 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill5.Append(gradientStopList5);
            gradientFill5.Append(linearGradientFill5);

            fillStyleList2.Append(solidFill77);
            fillStyleList2.Append(gradientFill4);
            fillStyleList2.Append(gradientFill5);

            A.LineStyleList lineStyleList2 = new A.LineStyleList();

            A.Outline outline17 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill78 = new A.SolidFill();
            A.SchemeColor schemeColor130 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill78.Append(schemeColor130);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter4 = new A.Miter() { Limit = 800000 };

            outline17.Append(solidFill78);
            outline17.Append(presetDash4);
            outline17.Append(miter4);

            A.Outline outline18 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill79 = new A.SolidFill();
            A.SchemeColor schemeColor131 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill79.Append(schemeColor131);
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter5 = new A.Miter() { Limit = 800000 };

            outline18.Append(solidFill79);
            outline18.Append(presetDash5);
            outline18.Append(miter5);

            A.Outline outline19 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill80 = new A.SolidFill();
            A.SchemeColor schemeColor132 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill80.Append(schemeColor132);
            A.PresetDash presetDash6 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter6 = new A.Miter() { Limit = 800000 };

            outline19.Append(solidFill80);
            outline19.Append(presetDash6);
            outline19.Append(miter6);

            lineStyleList2.Append(outline17);
            lineStyleList2.Append(outline18);
            lineStyleList2.Append(outline19);

            A.EffectStyleList effectStyleList2 = new A.EffectStyleList();

            A.EffectStyle effectStyle4 = new A.EffectStyle();
            A.EffectList effectList4 = new A.EffectList();

            effectStyle4.Append(effectList4);

            A.EffectStyle effectStyle5 = new A.EffectStyle();
            A.EffectList effectList5 = new A.EffectList();

            effectStyle5.Append(effectList5);

            A.EffectStyle effectStyle6 = new A.EffectStyle();

            A.EffectList effectList6 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex37 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha7 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex37.Append(alpha7);

            outerShadow2.Append(rgbColorModelHex37);

            effectList6.Append(outerShadow2);

            effectStyle6.Append(effectList6);

            effectStyleList2.Append(effectStyle4);
            effectStyleList2.Append(effectStyle5);
            effectStyleList2.Append(effectStyle6);

            A.BackgroundFillStyleList backgroundFillStyleList2 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill81 = new A.SolidFill();
            A.SchemeColor schemeColor133 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill81.Append(schemeColor133);

            A.SolidFill solidFill82 = new A.SolidFill();

            A.SchemeColor schemeColor134 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint12 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation17 = new A.SaturationModulation() { Val = 170000 };

            schemeColor134.Append(tint12);
            schemeColor134.Append(saturationModulation17);

            solidFill82.Append(schemeColor134);

            A.GradientFill gradientFill6 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList6 = new A.GradientStopList();

            A.GradientStop gradientStop16 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor135 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint13 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation18 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade14 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation20 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor135.Append(tint13);
            schemeColor135.Append(saturationModulation18);
            schemeColor135.Append(shade14);
            schemeColor135.Append(luminanceModulation20);

            gradientStop16.Append(schemeColor135);

            A.GradientStop gradientStop17 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor136 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint14 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation19 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade15 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation21 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor136.Append(tint14);
            schemeColor136.Append(saturationModulation19);
            schemeColor136.Append(shade15);
            schemeColor136.Append(luminanceModulation21);

            gradientStop17.Append(schemeColor136);

            A.GradientStop gradientStop18 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor137 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade16 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation20 = new A.SaturationModulation() { Val = 120000 };

            schemeColor137.Append(shade16);
            schemeColor137.Append(saturationModulation20);

            gradientStop18.Append(schemeColor137);

            gradientStopList6.Append(gradientStop16);
            gradientStopList6.Append(gradientStop17);
            gradientStopList6.Append(gradientStop18);
            A.LinearGradientFill linearGradientFill6 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill6.Append(gradientStopList6);
            gradientFill6.Append(linearGradientFill6);

            backgroundFillStyleList2.Append(solidFill81);
            backgroundFillStyleList2.Append(solidFill82);
            backgroundFillStyleList2.Append(gradientFill6);

            formatScheme2.Append(fillStyleList2);
            formatScheme2.Append(lineStyleList2);
            formatScheme2.Append(effectStyleList2);
            formatScheme2.Append(backgroundFillStyleList2);

            themeElements2.Append(colorScheme2);
            themeElements2.Append(fontScheme2);
            themeElements2.Append(formatScheme2);
            A.ObjectDefaults objectDefaults2 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList2 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList2 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension2 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily2 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily2.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension2.Append(themeFamily2);

            officeStyleSheetExtensionList2.Append(officeStyleSheetExtension2);

            theme2.Append(themeElements2);
            theme2.Append(objectDefaults2);
            theme2.Append(extraColorSchemeList2);
            theme2.Append(officeStyleSheetExtensionList2);

            themePart2.Theme = theme2;
        }

        // Generates content of slideLayoutPart9.
        private void GenerateSlideLayoutPart9Content(SlideLayoutPart slideLayoutPart9)
        {
            SlideLayout slideLayout9 = new SlideLayout() { Preserve = true, UserDrawn = true };
            slideLayout9.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout9.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout9.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData12 = new CommonSlideData() { Name = "Springboard" };

            ShapeTree shapeTree12 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties16 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties98 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties16 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties98 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties16.Append(nonVisualDrawingProperties98);
            nonVisualGroupShapeProperties16.Append(nonVisualGroupShapeDrawingProperties16);
            nonVisualGroupShapeProperties16.Append(applicationNonVisualDrawingProperties98);

            GroupShapeProperties groupShapeProperties16 = new GroupShapeProperties();

            A.TransformGroup transformGroup16 = new A.TransformGroup();
            A.Offset offset85 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents85 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset16 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents16 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup16.Append(offset85);
            transformGroup16.Append(extents85);
            transformGroup16.Append(childOffset16);
            transformGroup16.Append(childExtents16);

            groupShapeProperties16.Append(transformGroup16);

            Picture picture14 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties14 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties99 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties14 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks14 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties14.Append(pictureLocks14);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties99 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualPictureProperties14.Append(nonVisualDrawingProperties99);
            nonVisualPictureProperties14.Append(nonVisualPictureDrawingProperties14);
            nonVisualPictureProperties14.Append(applicationNonVisualDrawingProperties99);

            BlipFill blipFill14 = new BlipFill();

            A.Blip blip14 = new A.Blip() { Embed = "rId2" };

            A.BlipExtensionList blipExtensionList12 = new A.BlipExtensionList();

            A.BlipExtension blipExtension14 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi11 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi11.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension14.Append(useLocalDpi11);

            blipExtensionList12.Append(blipExtension14);

            blip14.Append(blipExtensionList12);

            A.Stretch stretch14 = new A.Stretch();
            A.FillRectangle fillRectangle12 = new A.FillRectangle();

            stretch14.Append(fillRectangle12);

            blipFill14.Append(blip14);
            blipFill14.Append(stretch14);

            ShapeProperties shapeProperties83 = new ShapeProperties();

            A.Transform2D transform2D70 = new A.Transform2D();
            A.Offset offset86 = new A.Offset() { X = -16772L, Y = 0L };
            A.Extents extents86 = new A.Extents() { Cx = 12208772L, Cy = 6858000L };

            transform2D70.Append(offset86);
            transform2D70.Append(extents86);

            A.PresetGeometry presetGeometry43 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList43 = new A.AdjustValueList();

            presetGeometry43.Append(adjustValueList43);

            shapeProperties83.Append(transform2D70);
            shapeProperties83.Append(presetGeometry43);

            picture14.Append(nonVisualPictureProperties14);
            picture14.Append(blipFill14);
            picture14.Append(shapeProperties83);

            Shape shape66 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties66 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties100 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 12" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties66 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks54 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties66.Append(shapeLocks54);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties100 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape54 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties100.Append(placeholderShape54);

            nonVisualShapeProperties66.Append(nonVisualDrawingProperties100);
            nonVisualShapeProperties66.Append(nonVisualShapeDrawingProperties66);
            nonVisualShapeProperties66.Append(applicationNonVisualDrawingProperties100);

            ShapeProperties shapeProperties84 = new ShapeProperties();

            A.Transform2D transform2D71 = new A.Transform2D();
            A.Offset offset87 = new A.Offset() { X = -17463L, Y = 0L };
            A.Extents extents87 = new A.Extents() { Cx = 12209463L, Cy = 6858000L };

            transform2D71.Append(offset87);
            transform2D71.Append(extents87);

            A.SolidFill solidFill83 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex38 = new A.RgbColorModelHex() { Val = "86D5AC" };
            A.Alpha alpha8 = new A.Alpha() { Val = 90000 };

            rgbColorModelHex38.Append(alpha8);

            solidFill83.Append(rgbColorModelHex38);

            shapeProperties84.Append(transform2D71);
            shapeProperties84.Append(solidFill83);

            TextBody textBody61 = new TextBody();

            A.BodyProperties bodyProperties61 = new A.BodyProperties();
            A.NormalAutoFit normalAutoFit5 = new A.NormalAutoFit();

            bodyProperties61.Append(normalAutoFit5);

            A.ListStyle listStyle61 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties37 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet23 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties73 = new A.DefaultRunProperties() { FontSize = 1400 };

            A.SolidFill solidFill84 = new A.SolidFill();

            A.SchemeColor schemeColor138 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.Alpha alpha9 = new A.Alpha() { Val = 0 };

            schemeColor138.Append(alpha9);

            solidFill84.Append(schemeColor138);

            defaultRunProperties73.Append(solidFill84);

            level1ParagraphProperties37.Append(noBullet23);
            level1ParagraphProperties37.Append(defaultRunProperties73);

            listStyle61.Append(level1ParagraphProperties37);

            A.Paragraph paragraph69 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties34 = new A.ParagraphProperties() { Level = 0 };

            A.Run run65 = new A.Run();
            A.RunProperties runProperties68 = new A.RunProperties() { Language = "en-US" };
            A.Text text67 = new A.Text();
            text67.Text = "Click to edit Master text styles";

            run65.Append(runProperties68);
            run65.Append(text67);

            paragraph69.Append(paragraphProperties34);
            paragraph69.Append(run65);

            textBody61.Append(bodyProperties61);
            textBody61.Append(listStyle61);
            textBody61.Append(paragraph69);

            shape66.Append(nonVisualShapeProperties66);
            shape66.Append(shapeProperties84);
            shape66.Append(textBody61);

            Shape shape67 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties67 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties101 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Picture Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties67 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks55 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties67.Append(shapeLocks55);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties101 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape55 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties101.Append(placeholderShape55);

            nonVisualShapeProperties67.Append(nonVisualDrawingProperties101);
            nonVisualShapeProperties67.Append(nonVisualShapeDrawingProperties67);
            nonVisualShapeProperties67.Append(applicationNonVisualDrawingProperties101);

            ShapeProperties shapeProperties85 = new ShapeProperties();

            A.Transform2D transform2D72 = new A.Transform2D();
            A.Offset offset88 = new A.Offset() { X = -1452284L, Y = -2022911L };
            A.Extents extents88 = new A.Extents() { Cx = 6120000L, Cy = 6120000L };

            transform2D72.Append(offset88);
            transform2D72.Append(extents88);

            A.PresetGeometry presetGeometry44 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList44 = new A.AdjustValueList();

            presetGeometry44.Append(adjustValueList44);

            A.SolidFill solidFill85 = new A.SolidFill();

            A.SchemeColor schemeColor139 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.LuminanceModulation luminanceModulation22 = new A.LuminanceModulation() { Val = 95000 };

            schemeColor139.Append(luminanceModulation22);

            solidFill85.Append(schemeColor139);

            shapeProperties85.Append(transform2D72);
            shapeProperties85.Append(presetGeometry44);
            shapeProperties85.Append(solidFill85);

            TextBody textBody62 = new TextBody();
            A.BodyProperties bodyProperties62 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle62 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties38 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet24 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties74 = new A.DefaultRunProperties();

            level1ParagraphProperties38.Append(noBullet24);
            level1ParagraphProperties38.Append(defaultRunProperties74);

            listStyle62.Append(level1ParagraphProperties38);

            A.Paragraph paragraph70 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties26 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph70.Append(endParagraphRunProperties26);

            textBody62.Append(bodyProperties62);
            textBody62.Append(listStyle62);
            textBody62.Append(paragraph70);

            shape67.Append(nonVisualShapeProperties67);
            shape67.Append(shapeProperties85);
            shape67.Append(textBody62);

            Shape shape68 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties68 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties102 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Title 8" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties68 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks56 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties68.Append(shapeLocks56);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties102 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape56 = new PlaceholderShape() { Type = PlaceholderValues.Title, HasCustomPrompt = true };

            applicationNonVisualDrawingProperties102.Append(placeholderShape56);

            nonVisualShapeProperties68.Append(nonVisualDrawingProperties102);
            nonVisualShapeProperties68.Append(nonVisualShapeDrawingProperties68);
            nonVisualShapeProperties68.Append(applicationNonVisualDrawingProperties102);

            ShapeProperties shapeProperties86 = new ShapeProperties();

            A.Transform2D transform2D73 = new A.Transform2D();
            A.Offset offset89 = new A.Offset() { X = -17463L, Y = 3870960L };
            A.Extents extents89 = new A.Extents() { Cx = 4203417L, Cy = 1441325L };

            transform2D73.Append(offset89);
            transform2D73.Append(extents89);

            shapeProperties86.Append(transform2D73);

            TextBody textBody63 = new TextBody();

            A.BodyProperties bodyProperties63 = new A.BodyProperties() { LeftInset = 684000, TopInset = 46800, RightInset = 144000, Anchor = A.TextAnchoringTypeValues.Bottom };
            A.NoAutoFit noAutoFit6 = new A.NoAutoFit();

            bodyProperties63.Append(noAutoFit6);

            A.ListStyle listStyle63 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties39 = new A.Level1ParagraphProperties();

            A.LineSpacing lineSpacing15 = new A.LineSpacing();
            A.SpacingPercent spacingPercent16 = new A.SpacingPercent() { Val = 80000 };

            lineSpacing15.Append(spacingPercent16);

            A.DefaultRunProperties defaultRunProperties75 = new A.DefaultRunProperties() { FontSize = 3600, Spacing = -150 };

            A.SolidFill solidFill86 = new A.SolidFill();
            A.SchemeColor schemeColor140 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill86.Append(schemeColor140);
            A.LatinFont latinFont64 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont64 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont64 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            defaultRunProperties75.Append(solidFill86);
            defaultRunProperties75.Append(latinFont64);
            defaultRunProperties75.Append(eastAsianFont64);
            defaultRunProperties75.Append(complexScriptFont64);

            level1ParagraphProperties39.Append(lineSpacing15);
            level1ParagraphProperties39.Append(defaultRunProperties75);

            listStyle63.Append(level1ParagraphProperties39);

            A.Paragraph paragraph71 = new A.Paragraph();

            A.Run run66 = new A.Run();
            A.RunProperties runProperties69 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text68 = new A.Text();
            text68.Text = "Name of springboard";

            run66.Append(runProperties69);
            run66.Append(text68);

            paragraph71.Append(run66);

            textBody63.Append(bodyProperties63);
            textBody63.Append(listStyle63);
            textBody63.Append(paragraph71);

            shape68.Append(nonVisualShapeProperties68);
            shape68.Append(shapeProperties86);
            shape68.Append(textBody63);

            Shape shape69 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties69 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties103 = new NonVisualDrawingProperties() { Id = (UInt32Value)11U, Name = "Text Placeholder 10" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties69 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks57 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties69.Append(shapeLocks57);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties103 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape57 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties103.Append(placeholderShape57);

            nonVisualShapeProperties69.Append(nonVisualDrawingProperties103);
            nonVisualShapeProperties69.Append(nonVisualShapeDrawingProperties69);
            nonVisualShapeProperties69.Append(applicationNonVisualDrawingProperties103);

            ShapeProperties shapeProperties87 = new ShapeProperties();

            A.Transform2D transform2D74 = new A.Transform2D();
            A.Offset offset90 = new A.Offset() { X = 0L, Y = 5476240L };
            A.Extents extents90 = new A.Extents() { Cx = 4186238L, Cy = 1381760L };

            transform2D74.Append(offset90);
            transform2D74.Append(extents90);

            shapeProperties87.Append(transform2D74);

            TextBody textBody64 = new TextBody();

            A.BodyProperties bodyProperties64 = new A.BodyProperties() { LeftInset = 684000, RightInset = 144000 };
            A.NoAutoFit noAutoFit7 = new A.NoAutoFit();

            bodyProperties64.Append(noAutoFit7);

            A.ListStyle listStyle64 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties40 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet25 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties76 = new A.DefaultRunProperties() { FontSize = 2000 };

            A.SolidFill solidFill87 = new A.SolidFill();
            A.SchemeColor schemeColor141 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill87.Append(schemeColor141);

            defaultRunProperties76.Append(solidFill87);

            level1ParagraphProperties40.Append(noBullet25);
            level1ParagraphProperties40.Append(defaultRunProperties76);

            listStyle64.Append(level1ParagraphProperties40);

            A.Paragraph paragraph72 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties35 = new A.ParagraphProperties() { Level = 0 };

            A.Run run67 = new A.Run();
            A.RunProperties runProperties70 = new A.RunProperties() { Language = "en-US", Dirty = false };
            A.Text text69 = new A.Text();
            text69.Text = "Click to edit Master text styles";

            run67.Append(runProperties70);
            run67.Append(text69);

            paragraph72.Append(paragraphProperties35);
            paragraph72.Append(run67);

            textBody64.Append(bodyProperties64);
            textBody64.Append(listStyle64);
            textBody64.Append(paragraph72);

            shape69.Append(nonVisualShapeProperties69);
            shape69.Append(shapeProperties87);
            shape69.Append(textBody64);

            shapeTree12.Append(nonVisualGroupShapeProperties16);
            shapeTree12.Append(groupShapeProperties16);
            shapeTree12.Append(picture14);
            shapeTree12.Append(shape66);
            shapeTree12.Append(shape67);
            shapeTree12.Append(shape68);
            shapeTree12.Append(shape69);

            CommonSlideDataExtensionList commonSlideDataExtensionList7 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension6 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId6 = new P14.CreationId() { Val = (UInt32Value)1732658620U };
            creationId6.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension6.Append(creationId6);

            commonSlideDataExtensionList7.Append(commonSlideDataExtension6);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList7);

            ColorMapOverride colorMapOverride10 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping10 = new A.MasterColorMapping();

            colorMapOverride10.Append(masterColorMapping10);

            slideLayout9.Append(commonSlideData12);
            slideLayout9.Append(colorMapOverride10);

            slideLayoutPart9.SlideLayout = slideLayout9;
        }

        #region Project Sources Slide.
        // Generates content of slidePart2.
        private void GenerateSlidePart2Content(SlidePart slidePart2)
        {
            Slide slide2 = new Slide();
            slide2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData13 = new CommonSlideData();

            ShapeTree shapeTree13 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties17 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties104 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties17 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties104 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties17.Append(nonVisualDrawingProperties104);
            nonVisualGroupShapeProperties17.Append(nonVisualGroupShapeDrawingProperties17);
            nonVisualGroupShapeProperties17.Append(applicationNonVisualDrawingProperties104);

            GroupShapeProperties groupShapeProperties17 = new GroupShapeProperties();

            A.TransformGroup transformGroup17 = new A.TransformGroup();
            A.Offset offset91 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents91 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset17 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents17 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup17.Append(offset91);
            transformGroup17.Append(extents91);
            transformGroup17.Append(childOffset17);
            transformGroup17.Append(childExtents17);

            groupShapeProperties17.Append(transformGroup17);

            Shape shape70 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties70 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties105 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties70 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks58 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties70.Append(shapeLocks58);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties105 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape58 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties105.Append(placeholderShape58);

            nonVisualShapeProperties70.Append(nonVisualDrawingProperties105);
            nonVisualShapeProperties70.Append(nonVisualShapeDrawingProperties70);
            nonVisualShapeProperties70.Append(applicationNonVisualDrawingProperties105);
            ShapeProperties shapeProperties88 = new ShapeProperties();

            TextBody textBody65 = new TextBody();
            A.BodyProperties bodyProperties65 = new A.BodyProperties();
            A.ListStyle listStyle65 = new A.ListStyle();

            A.Paragraph paragraph73 = new A.Paragraph();

            A.Run run68 = new A.Run();
            A.RunProperties runProperties71 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text70 = new A.Text();
            text70.Text = "Project Sources";

            run68.Append(runProperties71);
            run68.Append(text70);

            paragraph73.Append(run68);

            textBody65.Append(bodyProperties65);
            textBody65.Append(listStyle65);
            textBody65.Append(paragraph73);

            shape70.Append(nonVisualShapeProperties70);
            shape70.Append(shapeProperties88);
            shape70.Append(textBody65);

            Shape shape71 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties71 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties106 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "TextBox 3" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties71 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties106 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties71.Append(nonVisualDrawingProperties106);
            nonVisualShapeProperties71.Append(nonVisualShapeDrawingProperties71);
            nonVisualShapeProperties71.Append(applicationNonVisualDrawingProperties106);

            ShapeProperties shapeProperties89 = new ShapeProperties();

            A.Transform2D transform2D75 = new A.Transform2D();
            A.Offset offset92 = new A.Offset() { X = 864296L, Y = 1929008L };
            A.Extents extents92 = new A.Extents() { Cx = 9682619L, Cy = 1200329L };

            transform2D75.Append(offset92);
            transform2D75.Append(extents92);

            A.PresetGeometry presetGeometry45 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList45 = new A.AdjustValueList();

            presetGeometry45.Append(adjustValueList45);
            A.NoFill noFill19 = new A.NoFill();

            shapeProperties89.Append(transform2D75);
            shapeProperties89.Append(presetGeometry45);
            shapeProperties89.Append(noFill19);

            TextBody textBody66 = new TextBody();

            A.BodyProperties bodyProperties66 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit7 = new A.ShapeAutoFit();

            bodyProperties66.Append(shapeAutoFit7);
            A.ListStyle listStyle66 = new A.ListStyle();

            A.Paragraph paragraph74 = new A.Paragraph();

            A.Run run69 = new A.Run();
            A.RunProperties runProperties72 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text71 = new A.Text();
            text71.Text = "#0Source ";

            run69.Append(runProperties72);
            run69.Append(text71);

            A.Run run70 = new A.Run();
            A.RunProperties runProperties73 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text72 = new A.Text();
            text72.Text = _springboard.Project.Sources[0].Url; //"Url";

            run70.Append(runProperties73);
            run70.Append(text72);

            A.Run run71 = new A.Run();
            A.RunProperties runProperties74 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text73 = new A.Text();
            text73.Text = _springboard.Project.Sources[0].Area + ", " + _springboard.Project.Sources[0].Market; //", Area, Market";

            run71.Append(runProperties74);
            run71.Append(text73);

            paragraph74.Append(run69);
            paragraph74.Append(run70);
            paragraph74.Append(run71);

            A.Paragraph paragraph75 = new A.Paragraph();

            A.Run run72 = new A.Run();
            A.RunProperties runProperties75 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text74 = new A.Text();
            text74.Text = "#1Source";

            run72.Append(runProperties75);
            run72.Append(text74);

            A.Run run73 = new A.Run();
            A.RunProperties runProperties76 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text75 = new A.Text();
            text75.Text = "Url";

            run73.Append(runProperties76);
            run73.Append(text75);

            A.Run run74 = new A.Run();
            A.RunProperties runProperties77 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text76 = new A.Text();
            text76.Text = ", Area, Market";

            run74.Append(runProperties77);
            run74.Append(text76);

            paragraph75.Append(run72);
            paragraph75.Append(run73);
            paragraph75.Append(run74);

            A.Paragraph paragraph76 = new A.Paragraph();

            A.Run run75 = new A.Run();
            A.RunProperties runProperties78 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text77 = new A.Text();
            text77.Text = "#2Source";

            run75.Append(runProperties78);
            run75.Append(text77);

            A.Run run76 = new A.Run();
            A.RunProperties runProperties79 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text78 = new A.Text();
            text78.Text = "Url";

            run76.Append(runProperties79);
            run76.Append(text78);

            A.Run run77 = new A.Run();
            A.RunProperties runProperties80 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text79 = new A.Text();
            text79.Text = ", Area, Market";

            run77.Append(runProperties80);
            run77.Append(text79);

            paragraph76.Append(run75);
            paragraph76.Append(run76);
            paragraph76.Append(run77);

            A.Paragraph paragraph77 = new A.Paragraph();

            A.Run run78 = new A.Run();
            A.RunProperties runProperties81 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text80 = new A.Text();
            text80.Text = "…";

            run78.Append(runProperties81);
            run78.Append(text80);

            A.Run run79 = new A.Run();
            A.RunProperties runProperties82 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text81 = new A.Text();
            text81.Text = "etc";

            run79.Append(runProperties82);
            run79.Append(text81);
            A.EndParagraphRunProperties endParagraphRunProperties27 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            paragraph77.Append(run78);
            paragraph77.Append(run79);
            paragraph77.Append(endParagraphRunProperties27);

            textBody66.Append(bodyProperties66);
            textBody66.Append(listStyle66);
            textBody66.Append(paragraph74);
            textBody66.Append(paragraph75);
            textBody66.Append(paragraph76);
            textBody66.Append(paragraph77);

            shape71.Append(nonVisualShapeProperties71);
            shape71.Append(shapeProperties89);
            shape71.Append(textBody66);

            shapeTree13.Append(nonVisualGroupShapeProperties17);
            shapeTree13.Append(groupShapeProperties17);
            shapeTree13.Append(shape70);
            shapeTree13.Append(shape71);

            CommonSlideDataExtensionList commonSlideDataExtensionList8 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension7 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId7 = new P14.CreationId() { Val = (UInt32Value)1946748617U };
            creationId7.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension7.Append(creationId7);

            commonSlideDataExtensionList8.Append(commonSlideDataExtension7);

            commonSlideData13.Append(shapeTree13);
            commonSlideData13.Append(commonSlideDataExtensionList8);

            ColorMapOverride colorMapOverride11 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slide2.Append(commonSlideData13);
            slide2.Append(colorMapOverride11);

            slidePart2.Slide = slide2;
        }
        #endregion

        // Generates content of tableStylesPart1.
        private void GenerateTableStylesPart1Content(TableStylesPart tableStylesPart1)
        {
            A.TableStyleList tableStyleList1 = new A.TableStyleList() { Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
            tableStyleList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart1.TableStyleList = tableStyleList1;
        }

        #region Oval Slide
        // Generates content of slidePart3.
        private void GenerateSlidePart3Content(SlidePart slidePart3)
        {
            Slide slide3 = new Slide();
            slide3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData14 = new CommonSlideData();

            ShapeTree shapeTree14 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties18 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties107 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties18 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties107 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties18.Append(nonVisualDrawingProperties107);
            nonVisualGroupShapeProperties18.Append(nonVisualGroupShapeDrawingProperties18);
            nonVisualGroupShapeProperties18.Append(applicationNonVisualDrawingProperties107);

            GroupShapeProperties groupShapeProperties18 = new GroupShapeProperties();

            A.TransformGroup transformGroup18 = new A.TransformGroup();
            A.Offset offset93 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents93 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset18 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents18 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup18.Append(offset93);
            transformGroup18.Append(extents93);
            transformGroup18.Append(childOffset18);
            transformGroup18.Append(childExtents18);

            groupShapeProperties18.Append(transformGroup18);

            #region Project Question
            Shape shape72 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties72 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties108 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Title 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties72 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks59 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties72.Append(shapeLocks59);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties108 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape59 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties108.Append(placeholderShape59);

            nonVisualShapeProperties72.Append(nonVisualDrawingProperties108);
            nonVisualShapeProperties72.Append(nonVisualShapeDrawingProperties72);
            nonVisualShapeProperties72.Append(applicationNonVisualDrawingProperties108);

            ShapeProperties shapeProperties90 = new ShapeProperties();

            A.Transform2D transform2D76 = new A.Transform2D();
            A.Offset offset94 = new A.Offset() { X = 695325L, Y = 728663L };
            A.Extents extents94 = new A.Extents() { Cx = 10801350L, Cy = 2700336L };

            transform2D76.Append(offset94);
            transform2D76.Append(extents94);

            shapeProperties90.Append(transform2D76);

            TextBody textBody67 = new TextBody();
            A.BodyProperties bodyProperties67 = new A.BodyProperties();
            A.ListStyle listStyle67 = new A.ListStyle();

            A.Paragraph paragraph78 = new A.Paragraph();

            A.Run run80 = new A.Run();
            A.RunProperties runProperties83 = new A.RunProperties() { Language = "en-GB" };
            A.Text text82 = new A.Text();
            text82.Text = _springboard.Project.Question; //"$Question";

            run80.Append(runProperties83);
            run80.Append(text82);
            A.EndParagraphRunProperties endParagraphRunProperties28 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph78.Append(run80);
            paragraph78.Append(endParagraphRunProperties28);

            textBody67.Append(bodyProperties67);
            textBody67.Append(listStyle67);
            textBody67.Append(paragraph78);

            shape72.Append(nonVisualShapeProperties72);
            shape72.Append(shapeProperties90);
            shape72.Append(textBody67);
            #endregion

            #region Slide Intro Text.
            Shape shape73 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties73 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties109 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties73 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks60 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties73.Append(shapeLocks60);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties109 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape60 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties109.Append(placeholderShape60);

            nonVisualShapeProperties73.Append(nonVisualDrawingProperties109);
            nonVisualShapeProperties73.Append(nonVisualShapeDrawingProperties73);
            nonVisualShapeProperties73.Append(applicationNonVisualDrawingProperties109);

            ShapeProperties shapeProperties91 = new ShapeProperties();

            A.Transform2D transform2D77 = new A.Transform2D();
            A.Offset offset95 = new A.Offset() { X = 695325L, Y = 2859157L };
            A.Extents extents95 = new A.Extents() { Cx = 10801350L, Cy = 957470L };

            transform2D77.Append(offset95);
            transform2D77.Append(extents95);

            shapeProperties91.Append(transform2D77);

            TextBody textBody68 = new TextBody();
            A.BodyProperties bodyProperties68 = new A.BodyProperties();
            A.ListStyle listStyle68 = new A.ListStyle();

            A.Paragraph paragraph79 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties36 = new A.ParagraphProperties() { LeftMargin = 0, Indent = 0 };

            A.LineSpacing lineSpacing16 = new A.LineSpacing();
            A.SpacingPercent spacingPercent17 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing16.Append(spacingPercent17);
            A.NoBullet noBullet26 = new A.NoBullet();

            paragraphProperties36.Append(lineSpacing16);
            paragraphProperties36.Append(noBullet26);

            A.Run run81 = new A.Run();
            A.RunProperties runProperties84 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text83 = new A.Text();
            text83.Text = "We brought together rich & inspiring language ";

            run81.Append(runProperties84);
            run81.Append(text83);

            A.Run run82 = new A.Run();
            A.RunProperties runProperties85 = new A.RunProperties() { Language = "en-GB" };
            A.Text text84 = new A.Text();
            text84.Text = "from ";

            run82.Append(runProperties85);
            run82.Append(text84);

            A.Run run83 = new A.Run();

            A.RunProperties runProperties86 = new A.RunProperties() { Language = "en-GB" };

            A.SolidFill solidFill88 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex39 = new A.RgbColorModelHex() { Val = "FF0000" };

            //solidFill88.Append(rgbColorModelHex39);
            //runProperties86.Append(solidFill88);

            A.Text text85 = new A.Text();
            int sourcesPlusAreas = _springboard.Project.Sources.Length + _springboard.Project.Areas.Length;
            text85.Text = sourcesPlusAreas.ToString(); //"$SourceCount ";

            run83.Append(runProperties86);
            run83.Append(text85);

            /*
            A.Run run84 = new A.Run();
            A.RunProperties runProperties87 = new A.RunProperties() { Language = "en-GB" };
            A.Text text86 = new A.Text();
            text86.Text = "+ ";

            run84.Append(runProperties87);
            run84.Append(text86);

            A.Run run85 = new A.Run();

            A.RunProperties runProperties88 = new A.RunProperties() { Language = "en-GB" };

            A.SolidFill solidFill89 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex40 = new A.RgbColorModelHex() { Val = "FF0000" };

            solidFill89.Append(rgbColorModelHex40);

            runProperties88.Append(solidFill89);
            A.Text text87 = new A.Text();
            text87.Text = _springboard.Project.Areas.Length.ToString(); //"$AreaCount";

            run85.Append(runProperties88);
            run85.Append(text87);
            */

            A.Run run86 = new A.Run();
            A.RunProperties runProperties89 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text88 = new A.Text();
            text88.Text = " sources ";

            run86.Append(runProperties89);
            run86.Append(text88);

            A.Run run87 = new A.Run();
            A.RunProperties runProperties90 = new A.RunProperties() { Language = "en-GB" };
            A.Text text89 = new A.Text();
            text89.Text = "and from ";

            run87.Append(runProperties90);
            run87.Append(text89);

            A.Run run88 = new A.Run();

            A.RunProperties runProperties91 = new A.RunProperties() { Language = "en-GB" };

            A.SolidFill solidFill90 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex41 = new A.RgbColorModelHex() { Val = "FF0000" };

            //solidFill90.Append(rgbColorModelHex41);
            //runProperties91.Append(solidFill90);
            A.Text text90 = new A.Text();
            text90.Text = _springboard.Project.Markets.Length.ToString() + " markets."; //"$MarketCount";

            run88.Append(runProperties91);
            run88.Append(text90);

            A.EndParagraphRunProperties endParagraphRunProperties29 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            A.SolidFill solidFill91 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex42 = new A.RgbColorModelHex() { Val = "FF0000" };

            solidFill91.Append(rgbColorModelHex42);

            endParagraphRunProperties29.Append(solidFill91);

            paragraph79.Append(paragraphProperties36);
            paragraph79.Append(run81);
            paragraph79.Append(run82);
            paragraph79.Append(run83);
            //paragraph79.Append(run84);
            //paragraph79.Append(run85);
            paragraph79.Append(run86);
            paragraph79.Append(run87);
            paragraph79.Append(run88);
            paragraph79.Append(endParagraphRunProperties29);

            textBody68.Append(bodyProperties68);
            textBody68.Append(listStyle68);
            textBody68.Append(paragraph79);

            shape73.Append(nonVisualShapeProperties73);
            shape73.Append(shapeProperties91);
            shape73.Append(textBody68);

            GroupShape groupShape5 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties19 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties110 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Group 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList7 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension7 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement7 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{3BF156A0-D6C1-416A-ABFA-FBC1C94B6E11}\" />");

            nonVisualDrawingPropertiesExtension7.Append(openXmlUnknownElement7);

            nonVisualDrawingPropertiesExtensionList7.Append(nonVisualDrawingPropertiesExtension7);

            nonVisualDrawingProperties110.Append(nonVisualDrawingPropertiesExtensionList7);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties19 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties110 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties19.Append(nonVisualDrawingProperties110);
            nonVisualGroupShapeProperties19.Append(nonVisualGroupShapeDrawingProperties19);
            nonVisualGroupShapeProperties19.Append(applicationNonVisualDrawingProperties110);

            GroupShapeProperties groupShapeProperties19 = new GroupShapeProperties();

            A.TransformGroup transformGroup19 = new A.TransformGroup();
            A.Offset offset96 = new A.Offset() { X = 2913060L, Y = 3681877L };
            A.Extents extents96 = new A.Extents() { Cx = 6346014L, Cy = 2490501L };
            A.ChildOffset childOffset19 = new A.ChildOffset() { X = 1711308L, Y = 4066863L };
            A.ChildExtents childExtents19 = new A.ChildExtents() { Cx = 5365035L, Cy = 2105515L };

            transformGroup19.Append(offset96);
            transformGroup19.Append(extents96);
            transformGroup19.Append(childOffset19);
            transformGroup19.Append(childExtents19);

            A.SolidFill solidFill92 = new A.SolidFill();

            A.SchemeColor schemeColor142 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.LuminanceModulation luminanceModulation23 = new A.LuminanceModulation() { Val = 75000 };

            schemeColor142.Append(luminanceModulation23);

            solidFill92.Append(schemeColor142);

            groupShapeProperties19.Append(transformGroup19);
            groupShapeProperties19.Append(solidFill92);
            #endregion

            groupShape5.Append(nonVisualGroupShapeProperties19);
            groupShape5.Append(groupShapeProperties19);

            //A.Offset offset97 = new A.Offset() { X = 1711308L, Y = 4066863L };
            Shape shape74 = GenerateSlideShapeOval(_springboard.Project.Areas[0].Title, "Oval 5", 1711308L, 4066863L);
            //A.Offset offset98 = new A.Offset() { X = 2798571L, Y = 4066863L };
            Shape shape75 = GenerateSlideShapeOval(_springboard.Project.Areas[1].Title, "Oval 6", 2798571L, 4066863L);
            //A.Offset offset99 = new A.Offset() { X = 3885834L, Y = 4066863L };
            Shape shape76 = GenerateSlideShapeOval(_springboard.Project.Areas[2].Title, "Oval 7", 3885834L, 4066863L);
            //A.Offset offset100 = new A.Offset() { X = 4973097L, Y = 4066863L };
            Shape shape77 = GenerateSlideShapeOval(_springboard.Project.Areas[3].Title, "Oval 8", 4973097L, 4066863L);
            //A.Offset offset101 = new A.Offset() { X = 6060360L, Y = 4066863L };
            Shape shape78 = GenerateSlideShapeOval(_springboard.Project.Areas[4].Title, "Oval 9", 6060360L, 4066863L);
            // Top row.
            groupShape5.Append(shape74);
            groupShape5.Append(shape75);
            groupShape5.Append(shape76);
            groupShape5.Append(shape77);
            groupShape5.Append(shape78);

            Shape shape79 = GenerateSlideShapeOval(_springboard.Project.Areas[5].Title, "Oval 16", 1711308L, 5156395L);
            Shape shape80 = GenerateSlideShapeOval(_springboard.Project.Areas[6].Title, "Oval 17", 2798571L, 5156395L);
            Shape shape81 = GenerateSlideShapeOval(_springboard.Project.Areas[7].Title, "Oval 18", 3885834L, 5156395L);
            Shape shape82 = GenerateSlideShapeOval(_springboard.Project.Areas[8].Title, "Oval 19", 4973097L, 5156395L);
            //A.Offset offset106 = new A.Offset() { X = 6060360L, Y = 5156395L };
            Shape shape83 = GenerateSlideShapeOval(_springboard.Project.Areas[9].Title, "Oval 20", 6060360L, 5156395L);

            // Bottom row.
            groupShape5.Append(shape79);
            groupShape5.Append(shape80);
            groupShape5.Append(shape81);
            groupShape5.Append(shape82);
            groupShape5.Append(shape83);

            shapeTree14.Append(nonVisualGroupShapeProperties18);
            shapeTree14.Append(groupShapeProperties18);
            shapeTree14.Append(shape72);
            shapeTree14.Append(shape73);
            shapeTree14.Append(groupShape5);

            CommonSlideDataExtensionList commonSlideDataExtensionList9 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension8 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId8 = new P14.CreationId() { Val = (UInt32Value)46400944U };
            creationId8.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension8.Append(creationId8);

            commonSlideDataExtensionList9.Append(commonSlideDataExtension8);

            commonSlideData14.Append(shapeTree14);
            commonSlideData14.Append(commonSlideDataExtensionList9);

            ColorMapOverride colorMapOverride12 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping12 = new A.MasterColorMapping();

            colorMapOverride12.Append(masterColorMapping12);

            slide3.Append(commonSlideData14);
            slide3.Append(colorMapOverride12);

            slidePart3.Slide = slide3;
        }

        private Shape GenerateSlideShapeOval(string areaTitle, string nonVisualDrawingName, long xOffset, long yOffset)
        {
            Shape shape75 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties75 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties112 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = nonVisualDrawingName };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties75 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties112 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties75.Append(nonVisualDrawingProperties112);
            nonVisualShapeProperties75.Append(nonVisualShapeDrawingProperties75);
            nonVisualShapeProperties75.Append(applicationNonVisualDrawingProperties112);

            ShapeProperties shapeProperties93 = new ShapeProperties();

            A.Transform2D transform2D79 = new A.Transform2D();
            A.Offset offset98 = new A.Offset() { X = xOffset, Y = yOffset };
            A.Extents extents98 = new A.Extents() { Cx = 1015983L, Cy = 1015983L };

            transform2D79.Append(offset98);
            transform2D79.Append(extents98);

            A.PresetGeometry presetGeometry47 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList47 = new A.AdjustValueList();

            presetGeometry47.Append(adjustValueList47);
            A.GroupFill groupFill2 = new A.GroupFill();

            A.Outline outline21 = new A.Outline();
            A.NoFill noFill21 = new A.NoFill();

            outline21.Append(noFill21);

            shapeProperties93.Append(transform2D79);
            shapeProperties93.Append(presetGeometry47);
            shapeProperties93.Append(groupFill2);
            shapeProperties93.Append(outline21);

            ShapeStyle shapeStyle12 = new ShapeStyle();

            A.LineReference lineReference12 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor147 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade18 = new A.Shade() { Val = 50000 };

            schemeColor147.Append(shade18);

            lineReference12.Append(schemeColor147);

            A.FillReference fillReference12 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor148 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference12.Append(schemeColor148);

            A.EffectReference effectReference12 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor149 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference12.Append(schemeColor149);

            A.FontReference fontReference12 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor150 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference12.Append(schemeColor150);

            shapeStyle12.Append(lineReference12);
            shapeStyle12.Append(fillReference12);
            shapeStyle12.Append(effectReference12);
            shapeStyle12.Append(fontReference12);

            TextBody textBody70 = new TextBody();
            A.BodyProperties bodyProperties70 = new A.BodyProperties() { LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle70 = new A.ListStyle();

            A.Paragraph paragraph82 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties39 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.Run run91 = new A.Run();

            A.RunProperties runProperties94 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Dirty = false };
            A.LatinFont latinFont68 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties94.Append(latinFont68);
            A.Text text93 = new A.Text();
            text93.Text = "#1";

            run91.Append(runProperties94);
            run91.Append(text93);

            paragraph82.Append(paragraphProperties39);
            paragraph82.Append(run91);

            A.Paragraph paragraph83 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties40 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.Run run92 = new A.Run();

            A.RunProperties runProperties95 = new A.RunProperties() { Language = "en-US", FontSize = 1200 };
            A.LatinFont latinFont69 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties95.Append(latinFont69);
            A.Text text94 = new A.Text();
            text94.Text = areaTitle;// "Area$Title";

            run92.Append(runProperties95);
            run92.Append(text94);

            A.EndParagraphRunProperties endParagraphRunProperties31 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Dirty = false };
            A.LatinFont latinFont70 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties31.Append(latinFont70);

            paragraph83.Append(paragraphProperties40);
            paragraph83.Append(run92);
            paragraph83.Append(endParagraphRunProperties31);

            textBody70.Append(bodyProperties70);
            textBody70.Append(listStyle70);
            textBody70.Append(paragraph82);
            textBody70.Append(paragraph83);

            shape75.Append(nonVisualShapeProperties75);
            shape75.Append(shapeProperties93);
            shape75.Append(shapeStyle12);
            shape75.Append(textBody70);

            return shape75;
        }
        #endregion

        // Generates content of slidePart4.
        private void GenerateSlidePart4Content(SlidePart slidePart4)
        {
            Slide slide4 = new Slide();
            slide4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData15 = new CommonSlideData();

            ShapeTree shapeTree15 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties20 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties121 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties20 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties121 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties20.Append(nonVisualDrawingProperties121);
            nonVisualGroupShapeProperties20.Append(nonVisualGroupShapeDrawingProperties20);
            nonVisualGroupShapeProperties20.Append(applicationNonVisualDrawingProperties121);

            GroupShapeProperties groupShapeProperties20 = new GroupShapeProperties();

            A.TransformGroup transformGroup20 = new A.TransformGroup();
            A.Offset offset107 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents107 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset20 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents20 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup20.Append(offset107);
            transformGroup20.Append(extents107);
            transformGroup20.Append(childOffset20);
            transformGroup20.Append(childExtents20);

            groupShapeProperties20.Append(transformGroup20);

            Shape shape84 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties84 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties122 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties84 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks61 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties84.Append(shapeLocks61);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties122 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape61 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties122.Append(placeholderShape61);

            nonVisualShapeProperties84.Append(nonVisualDrawingProperties122);
            nonVisualShapeProperties84.Append(nonVisualShapeDrawingProperties84);
            nonVisualShapeProperties84.Append(applicationNonVisualDrawingProperties122);
            ShapeProperties shapeProperties102 = new ShapeProperties();

            TextBody textBody79 = new TextBody();
            A.BodyProperties bodyProperties79 = new A.BodyProperties();
            A.ListStyle listStyle79 = new A.ListStyle();

            A.Paragraph paragraph100 = new A.Paragraph();

            A.Run run109 = new A.Run();
            A.RunProperties runProperties112 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text111 = new A.Text();
            text111.Text = "$";

            run109.Append(runProperties112);
            run109.Append(text111);

            A.Run run110 = new A.Run();
            A.RunProperties runProperties113 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text112 = new A.Text();
            text112.Text = _springboard.Project.WordLists[0].Title;//"WordList.Title";

            run110.Append(runProperties113);
            run110.Append(text112);
            A.EndParagraphRunProperties endParagraphRunProperties40 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            paragraph100.Append(run109);
            paragraph100.Append(run110);
            paragraph100.Append(endParagraphRunProperties40);

            textBody79.Append(bodyProperties79);
            textBody79.Append(listStyle79);
            textBody79.Append(paragraph100);

            shape84.Append(nonVisualShapeProperties84);
            shape84.Append(shapeProperties102);
            shape84.Append(textBody79);

            Shape shape85 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties85 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties123 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "TextBox 3" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties85 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties123 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties85.Append(nonVisualDrawingProperties123);
            nonVisualShapeProperties85.Append(nonVisualShapeDrawingProperties85);
            nonVisualShapeProperties85.Append(applicationNonVisualDrawingProperties123);

            ShapeProperties shapeProperties103 = new ShapeProperties();

            A.Transform2D transform2D88 = new A.Transform2D();
            A.Offset offset108 = new A.Offset() { X = 1002082L, Y = 2029216L };
            A.Extents extents108 = new A.Extents() { Cx = 9682619L, Cy = 1200329L };

            transform2D88.Append(offset108);
            transform2D88.Append(extents108);

            A.PresetGeometry presetGeometry56 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList56 = new A.AdjustValueList();

            presetGeometry56.Append(adjustValueList56);
            A.NoFill noFill30 = new A.NoFill();

            shapeProperties103.Append(transform2D88);
            shapeProperties103.Append(presetGeometry56);
            shapeProperties103.Append(noFill30);

            TextBody textBody80 = new TextBody();

            A.BodyProperties bodyProperties80 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit8 = new A.ShapeAutoFit();

            bodyProperties80.Append(shapeAutoFit8);
            A.ListStyle listStyle80 = new A.ListStyle();

            A.Paragraph paragraph101 = new A.Paragraph();

            A.Run run111 = new A.Run();
            A.RunProperties runProperties114 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text113 = new A.Text();
            text113.Text = _springboard.Project.WordLists[0].Words[0];//"#0Word";

            run111.Append(runProperties114);
            run111.Append(text113);

            paragraph101.Append(run111);

            A.Paragraph paragraph102 = new A.Paragraph();

            A.Run run112 = new A.Run();
            A.RunProperties runProperties115 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text114 = new A.Text();
            text114.Text = "#1Word";

            run112.Append(runProperties115);
            run112.Append(text114);

            paragraph102.Append(run112);

            A.Paragraph paragraph103 = new A.Paragraph();

            A.Run run113 = new A.Run();
            A.RunProperties runProperties116 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text115 = new A.Text();
            text115.Text = "#2Word";

            run113.Append(runProperties116);
            run113.Append(text115);

            paragraph103.Append(run113);

            A.Paragraph paragraph104 = new A.Paragraph();

            A.Run run114 = new A.Run();
            A.RunProperties runProperties117 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text116 = new A.Text();
            text116.Text = "…";

            run114.Append(runProperties117);
            run114.Append(text116);

            A.Run run115 = new A.Run();
            A.RunProperties runProperties118 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text117 = new A.Text();
            text117.Text = "etc";

            run115.Append(runProperties118);
            run115.Append(text117);
            A.EndParagraphRunProperties endParagraphRunProperties41 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            paragraph104.Append(run114);
            paragraph104.Append(run115);
            paragraph104.Append(endParagraphRunProperties41);

            textBody80.Append(bodyProperties80);
            textBody80.Append(listStyle80);
            textBody80.Append(paragraph101);
            textBody80.Append(paragraph102);
            textBody80.Append(paragraph103);
            textBody80.Append(paragraph104);

            shape85.Append(nonVisualShapeProperties85);
            shape85.Append(shapeProperties103);
            shape85.Append(textBody80);

            shapeTree15.Append(nonVisualGroupShapeProperties20);
            shapeTree15.Append(groupShapeProperties20);
            shapeTree15.Append(shape84);
            shapeTree15.Append(shape85);

            CommonSlideDataExtensionList commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension9 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId9 = new P14.CreationId() { Val = (UInt32Value)1084620351U };
            creationId9.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension9.Append(creationId9);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension9);

            commonSlideData15.Append(shapeTree15);
            commonSlideData15.Append(commonSlideDataExtensionList10);

            ColorMapOverride colorMapOverride13 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping13 = new A.MasterColorMapping();

            colorMapOverride13.Append(masterColorMapping13);

            slide4.Append(commonSlideData15);
            slide4.Append(colorMapOverride13);

            slidePart4.Slide = slide4;
        }

        // Generates content of slidePart5.
        private void GenerateSlidePart5Content(SlidePart slidePart5)
        {
            Slide slide5 = new Slide();
            slide5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData16 = new CommonSlideData();

            ShapeTree shapeTree16 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties21 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties124 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties21 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties124 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties21.Append(nonVisualDrawingProperties124);
            nonVisualGroupShapeProperties21.Append(nonVisualGroupShapeDrawingProperties21);
            nonVisualGroupShapeProperties21.Append(applicationNonVisualDrawingProperties124);

            GroupShapeProperties groupShapeProperties21 = new GroupShapeProperties();

            A.TransformGroup transformGroup21 = new A.TransformGroup();
            A.Offset offset109 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents109 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset21 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents21 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup21.Append(offset109);
            transformGroup21.Append(extents109);
            transformGroup21.Append(childOffset21);
            transformGroup21.Append(childExtents21);

            groupShapeProperties21.Append(transformGroup21);

            Shape shape86 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties86 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties125 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties86 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks62 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties86.Append(shapeLocks62);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties125 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape62 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties125.Append(placeholderShape62);

            nonVisualShapeProperties86.Append(nonVisualDrawingProperties125);
            nonVisualShapeProperties86.Append(nonVisualShapeDrawingProperties86);
            nonVisualShapeProperties86.Append(applicationNonVisualDrawingProperties125);
            ShapeProperties shapeProperties104 = new ShapeProperties();

            TextBody textBody81 = new TextBody();
            A.BodyProperties bodyProperties81 = new A.BodyProperties();
            A.ListStyle listStyle81 = new A.ListStyle();

            A.Paragraph paragraph105 = new A.Paragraph();

            A.Run run116 = new A.Run();
            A.RunProperties runProperties119 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text118 = new A.Text();
            text118.Text = "$";

            run116.Append(runProperties119);
            run116.Append(text118);

            A.Run run117 = new A.Run();
            A.RunProperties runProperties120 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text119 = new A.Text();
            text119.Text = _springboard.Project.WordClouds[0].Title;//"WordCloud.Title";

            run117.Append(runProperties120);
            run117.Append(text119);
            A.EndParagraphRunProperties endParagraphRunProperties42 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            paragraph105.Append(run116);
            paragraph105.Append(run117);
            paragraph105.Append(endParagraphRunProperties42);

            textBody81.Append(bodyProperties81);
            textBody81.Append(listStyle81);
            textBody81.Append(paragraph105);

            shape86.Append(nonVisualShapeProperties86);
            shape86.Append(shapeProperties104);
            shape86.Append(textBody81);

            Shape shape87 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties87 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties126 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties87 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks63 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties87.Append(shapeLocks63);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties126 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape63 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties126.Append(placeholderShape63);

            nonVisualShapeProperties87.Append(nonVisualDrawingProperties126);
            nonVisualShapeProperties87.Append(nonVisualShapeDrawingProperties87);
            nonVisualShapeProperties87.Append(applicationNonVisualDrawingProperties126);

            ShapeProperties shapeProperties105 = new ShapeProperties();

            A.Transform2D transform2D89 = new A.Transform2D();
            A.Offset offset110 = new A.Offset() { X = 1271523L, Y = 1568120L };
            A.Extents extents110 = new A.Extents() { Cx = 9814012L, Cy = 4870258L };

            transform2D89.Append(offset110);
            transform2D89.Append(extents110);

            shapeProperties105.Append(transform2D89);

            shape87.Append(nonVisualShapeProperties87);
            shape87.Append(shapeProperties105);

            Shape shape88 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties88 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties127 = new NonVisualDrawingProperties() { Id = (UInt32Value)19U, Name = "TextBox 18" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties88 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties127 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties88.Append(nonVisualDrawingProperties127);
            nonVisualShapeProperties88.Append(nonVisualShapeDrawingProperties88);
            nonVisualShapeProperties88.Append(applicationNonVisualDrawingProperties127);

            ShapeProperties shapeProperties106 = new ShapeProperties();

            A.Transform2D transform2D90 = new A.Transform2D();
            A.Offset offset111 = new A.Offset() { X = 4647156L, Y = 4183693L };
            A.Extents extents111 = new A.Extents() { Cx = 3356976L, Cy = 369332L };

            transform2D90.Append(offset111);
            transform2D90.Append(extents111);

            A.PresetGeometry presetGeometry57 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList57 = new A.AdjustValueList();

            presetGeometry57.Append(adjustValueList57);
            A.NoFill noFill31 = new A.NoFill();

            shapeProperties106.Append(transform2D90);
            shapeProperties106.Append(presetGeometry57);
            shapeProperties106.Append(noFill31);

            TextBody textBody82 = new TextBody();

            A.BodyProperties bodyProperties82 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit9 = new A.ShapeAutoFit();

            bodyProperties82.Append(shapeAutoFit9);
            A.ListStyle listStyle82 = new A.ListStyle();

            A.Paragraph paragraph106 = new A.Paragraph();

            A.Run run118 = new A.Run();
            A.RunProperties runProperties121 = new A.RunProperties() { Language = "en-GB", Dirty = false };
            A.Text text120 = new A.Text();
            text120.Text = "Image from ";

            run118.Append(runProperties121);
            run118.Append(text120);

            A.Run run119 = new A.Run();
            A.RunProperties runProperties122 = new A.RunProperties() { Language = "en-GB", Dirty = false, SpellingError = true };
            A.Text text121 = new A.Text();
            text121.Text = "Url";

            run119.Append(runProperties122);
            run119.Append(text121);
            A.EndParagraphRunProperties endParagraphRunProperties43 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            paragraph106.Append(run118);
            paragraph106.Append(run119);
            paragraph106.Append(endParagraphRunProperties43);

            textBody82.Append(bodyProperties82);
            textBody82.Append(listStyle82);
            textBody82.Append(paragraph106);

            shape88.Append(nonVisualShapeProperties88);
            shape88.Append(shapeProperties106);
            shape88.Append(textBody82);

            shapeTree16.Append(nonVisualGroupShapeProperties21);
            shapeTree16.Append(groupShapeProperties21);
            shapeTree16.Append(shape86);
            shapeTree16.Append(shape87);
            shapeTree16.Append(shape88);

            CommonSlideDataExtensionList commonSlideDataExtensionList11 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension10 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId10 = new P14.CreationId() { Val = (UInt32Value)1059833608U };
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList11.Append(commonSlideDataExtension10);

            commonSlideData16.Append(shapeTree16);
            commonSlideData16.Append(commonSlideDataExtensionList11);

            ColorMapOverride colorMapOverride14 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping14 = new A.MasterColorMapping();

            colorMapOverride14.Append(masterColorMapping14);

            slide5.Append(commonSlideData16);
            slide5.Append(colorMapOverride14);

            slidePart5.Slide = slide5;
        }

        // Generates content of viewPropertiesPart1.
        private void GenerateViewPropertiesPart1Content(ViewPropertiesPart viewPropertiesPart1)
        {
            ViewProperties viewProperties1 = new ViewProperties() { LastView = ViewValues.SlideThumbnailView };
            viewProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            viewProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            viewProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            NormalViewProperties normalViewProperties1 = new NormalViewProperties() { HorizontalBarState = SplitterBarStateValues.Maximized };
            RestoredLeft restoredLeft1 = new RestoredLeft() { Size = 15014, AutoAdjust = false };
            RestoredTop restoredTop1 = new RestoredTop() { Size = 95665 };

            normalViewProperties1.Append(restoredLeft1);
            normalViewProperties1.Append(restoredTop1);

            SlideViewProperties slideViewProperties1 = new SlideViewProperties();

            CommonSlideViewProperties commonSlideViewProperties1 = new CommonSlideViewProperties() { SnapToGrid = false, SnapToObjects = true };

            CommonViewProperties commonViewProperties1 = new CommonViewProperties() { VariableScale = true };

            ScaleFactor scaleFactor1 = new ScaleFactor();
            A.ScaleX scaleX1 = new A.ScaleX() { Numerator = 76, Denominator = 100 };
            A.ScaleY scaleY1 = new A.ScaleY() { Numerator = 76, Denominator = 100 };

            scaleFactor1.Append(scaleX1);
            scaleFactor1.Append(scaleY1);
            Origin origin1 = new Origin() { X = 720L, Y = 96L };

            commonViewProperties1.Append(scaleFactor1);
            commonViewProperties1.Append(origin1);

            GuideList guideList1 = new GuideList();
            Guide guide1 = new Guide() { Orientation = DirectionValues.Horizontal, Position = 2160 };
            Guide guide2 = new Guide() { Position = 3840 };

            guideList1.Append(guide1);
            guideList1.Append(guide2);

            commonSlideViewProperties1.Append(commonViewProperties1);
            commonSlideViewProperties1.Append(guideList1);

            slideViewProperties1.Append(commonSlideViewProperties1);

            NotesTextViewProperties notesTextViewProperties1 = new NotesTextViewProperties();

            CommonViewProperties commonViewProperties2 = new CommonViewProperties();

            ScaleFactor scaleFactor2 = new ScaleFactor();
            A.ScaleX scaleX2 = new A.ScaleX() { Numerator = 1, Denominator = 1 };
            A.ScaleY scaleY2 = new A.ScaleY() { Numerator = 1, Denominator = 1 };

            scaleFactor2.Append(scaleX2);
            scaleFactor2.Append(scaleY2);
            Origin origin2 = new Origin() { X = 0L, Y = 0L };

            commonViewProperties2.Append(scaleFactor2);
            commonViewProperties2.Append(origin2);

            notesTextViewProperties1.Append(commonViewProperties2);
            GridSpacing gridSpacing1 = new GridSpacing() { Cx = 72008L, Cy = 72008L };

            viewProperties1.Append(normalViewProperties1);
            viewProperties1.Append(slideViewProperties1);
            viewProperties1.Append(notesTextViewProperties1);
            viewProperties1.Append(gridSpacing1);

            viewPropertiesPart1.ViewProperties = viewProperties1;
        }

        #region Springboard Themes & Scattered Ovals Slide.
        // Generates content of slidePart6.
        private void GenerateSlidePart6Content(SlidePart slidePart6)
        {
            Slide slide6 = new Slide();
            slide6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData17 = new CommonSlideData();

            ShapeTree shapeTree17 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties22 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties128 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties22 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties128 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties22.Append(nonVisualDrawingProperties128);
            nonVisualGroupShapeProperties22.Append(nonVisualGroupShapeDrawingProperties22);
            nonVisualGroupShapeProperties22.Append(applicationNonVisualDrawingProperties128);

            GroupShapeProperties groupShapeProperties22 = new GroupShapeProperties();

            A.TransformGroup transformGroup22 = new A.TransformGroup();
            A.Offset offset112 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents112 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset22 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents22 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup22.Append(offset112);
            transformGroup22.Append(extents112);
            transformGroup22.Append(childOffset22);
            transformGroup22.Append(childExtents22);

            groupShapeProperties22.Append(transformGroup22);

            #region Placeholder Shape
            Shape shape89 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties89 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties129 = new NonVisualDrawingProperties() { Id = (UInt32Value)29U, Name = "Text Placeholder 28" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties89 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks64 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties89.Append(shapeLocks64);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties129 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape64 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties129.Append(placeholderShape64);

            nonVisualShapeProperties89.Append(nonVisualDrawingProperties129);
            nonVisualShapeProperties89.Append(nonVisualShapeDrawingProperties89);
            nonVisualShapeProperties89.Append(applicationNonVisualDrawingProperties129);

            ShapeProperties shapeProperties107 = new ShapeProperties();

            A.SolidFill solidFill93 = new A.SolidFill();

            A.SchemeColor schemeColor183 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.LuminanceModulation luminanceModulation24 = new A.LuminanceModulation() { Val = 60000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 40000 };
            A.Alpha alpha10 = new A.Alpha() { Val = 90000 };

            schemeColor183.Append(luminanceModulation24);
            schemeColor183.Append(luminanceOffset1);
            schemeColor183.Append(alpha10);

            solidFill93.Append(schemeColor183);

            shapeProperties107.Append(solidFill93);

            TextBody textBody83 = new TextBody();
            A.BodyProperties bodyProperties83 = new A.BodyProperties();
            A.ListStyle listStyle83 = new A.ListStyle();

            A.Paragraph paragraph107 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties44 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph107.Append(endParagraphRunProperties44);

            textBody83.Append(bodyProperties83);
            textBody83.Append(listStyle83);
            textBody83.Append(paragraph107);

            shape89.Append(nonVisualShapeProperties89);
            shape89.Append(shapeProperties107);
            shape89.Append(textBody83);
            #endregion

            #region Title Shape
            Shape shape90 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties90 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties130 = new NonVisualDrawingProperties() { Id = (UInt32Value)20U, Name = "Title 19" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties90 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks65 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties90.Append(shapeLocks65);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties130 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape65 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties130.Append(placeholderShape65);

            nonVisualShapeProperties90.Append(nonVisualDrawingProperties130);
            nonVisualShapeProperties90.Append(nonVisualShapeDrawingProperties90);
            nonVisualShapeProperties90.Append(applicationNonVisualDrawingProperties130);
            ShapeProperties shapeProperties108 = new ShapeProperties();

            TextBody textBody84 = new TextBody();
            A.BodyProperties bodyProperties84 = new A.BodyProperties();
            A.ListStyle listStyle84 = new A.ListStyle();

            A.Paragraph paragraph108 = new A.Paragraph();

            A.Run run120 = new A.Run();
            A.RunProperties runProperties123 = new A.RunProperties() { Language = "en-US" };
            A.Text text122 = new A.Text();
            text122.Text = _springboard.Project.Areas[0].Springboards[0].Title; //"$SpringBoard.Title";

            run120.Append(runProperties123);
            run120.Append(text122);
            A.EndParagraphRunProperties endParagraphRunProperties45 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph108.Append(run120);
            paragraph108.Append(endParagraphRunProperties45);

            textBody84.Append(bodyProperties84);
            textBody84.Append(listStyle84);
            textBody84.Append(paragraph108);

            shape90.Append(nonVisualShapeProperties90);
            shape90.Append(shapeProperties108);
            shape90.Append(textBody84);
            #endregion

            #region Placeholder Text Shape
            Shape shape91 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties91 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties131 = new NonVisualDrawingProperties() { Id = (UInt32Value)25U, Name = "Text Placeholder 24" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties91 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks66 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties91.Append(shapeLocks66);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties131 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape66 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties131.Append(placeholderShape66);

            nonVisualShapeProperties91.Append(nonVisualDrawingProperties131);
            nonVisualShapeProperties91.Append(nonVisualShapeDrawingProperties91);
            nonVisualShapeProperties91.Append(applicationNonVisualDrawingProperties131);

            ShapeProperties shapeProperties109 = new ShapeProperties();

            A.Transform2D transform2D91 = new A.Transform2D();
            A.Offset offset113 = new A.Offset() { X = 0L, Y = 5308480L };
            A.Extents extents113 = new A.Extents() { Cx = 4186238L, Cy = 991950L };

            transform2D91.Append(offset113);
            transform2D91.Append(extents113);

            shapeProperties109.Append(transform2D91);

            TextBody textBody85 = new TextBody();
            A.BodyProperties bodyProperties85 = new A.BodyProperties();
            A.ListStyle listStyle85 = new A.ListStyle();

            A.Paragraph paragraph109 = new A.Paragraph();

            A.Run run121 = new A.Run();
            A.RunProperties runProperties124 = new A.RunProperties() { Language = "en-US" };
            A.Text text123 = new A.Text();
            text123.Text = _springboard.Project.Areas[0].Springboards[0].Description; //"$Springboard.Description";

            run121.Append(runProperties124);
            run121.Append(text123);
            A.EndParagraphRunProperties endParagraphRunProperties46 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph109.Append(run121);
            paragraph109.Append(endParagraphRunProperties46);

            textBody85.Append(bodyProperties85);
            textBody85.Append(listStyle85);
            textBody85.Append(paragraph109);

            shape91.Append(nonVisualShapeProperties91);
            shape91.Append(shapeProperties109);
            shape91.Append(textBody85);

            Picture picture15 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties15 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties132 = new NonVisualDrawingProperties() { Id = (UInt32Value)32U, Name = "Picture 31" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties15 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks15 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties15.Append(pictureLocks15);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties132 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties15.Append(nonVisualDrawingProperties132);
            nonVisualPictureProperties15.Append(nonVisualPictureDrawingProperties15);
            nonVisualPictureProperties15.Append(applicationNonVisualDrawingProperties132);

            BlipFill blipFill15 = new BlipFill();

            A.Blip blip15 = new A.Blip() { Embed = "rId3" };

            A.BlipExtensionList blipExtensionList13 = new A.BlipExtensionList();

            A.BlipExtension blipExtension15 = new A.BlipExtension() { Uri = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}" };

            A14.ImageProperties imageProperties4 = new A14.ImageProperties();
            imageProperties4.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A14.ImageLayer imageLayer4 = new A14.ImageLayer() { Embed = "rId4" };

            A14.ImageEffect imageEffect4 = new A14.ImageEffect();
            A14.BrightnessContrast brightnessContrast4 = new A14.BrightnessContrast() { Bright = 100000 };

            imageEffect4.Append(brightnessContrast4);

            imageLayer4.Append(imageEffect4);

            imageProperties4.Append(imageLayer4);

            blipExtension15.Append(imageProperties4);

            A.BlipExtension blipExtension16 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi12 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi12.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension16.Append(useLocalDpi12);

            blipExtensionList13.Append(blipExtension15);
            blipExtensionList13.Append(blipExtension16);

            blip15.Append(blipExtensionList13);

            A.Stretch stretch15 = new A.Stretch();
            A.FillRectangle fillRectangle13 = new A.FillRectangle();

            stretch15.Append(fillRectangle13);

            blipFill15.Append(blip15);
            blipFill15.Append(stretch15);

            ShapeProperties shapeProperties110 = new ShapeProperties();

            A.Transform2D transform2D92 = new A.Transform2D();
            A.Offset offset114 = new A.Offset() { X = 10919356L, Y = 6465900L };
            A.Extents extents114 = new A.Extents() { Cx = 1095427L, Cy = 260968L };

            transform2D92.Append(offset114);
            transform2D92.Append(extents114);

            A.PresetGeometry presetGeometry58 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList58 = new A.AdjustValueList();

            presetGeometry58.Append(adjustValueList58);

            shapeProperties110.Append(transform2D92);
            shapeProperties110.Append(presetGeometry58);

            picture15.Append(nonVisualPictureProperties15);
            picture15.Append(blipFill15);
            picture15.Append(shapeProperties110);
            #endregion

            #region Oval Shape, Theme Title, Text, Source Url #6
            Shape shape92 = GenerateSpringboardDetailThemeOffset(19U, "Oval 18", 7727716L, 3731327L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[6].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[6].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[6].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #1
            Shape shape93 = GenerateSpringboardDetailThemeOffset(21U, "Oval 20", 4185956L, 3459937L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[1].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[1].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[1].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #2
            Shape shape94 = GenerateSpringboardDetailThemeOffset(22U, "Oval 21", 6292005L, 524151L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[2].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[2].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[2].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #9
            Shape shape95 = GenerateSpringboardDetailThemeOffset(24U, "Oval 23", 9671717L, 4390769L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[9].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[9].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[9].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #4
            Shape shape96 = GenerateSpringboardDetailThemeOffset(23U, "Oval 22", 5903821L, 4669555L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[4].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[4].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[4].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #7
            Shape shape97 = GenerateSpringboardDetailThemeOffset(26U, "Oval 25", 9544250L, 262305L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[7].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[7].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[7].SourceUrl);
            /*
            Shape shape97 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties97 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties138 = new NonVisualDrawingProperties() { Id = (UInt32Value)26U, Name = "Oval 25" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties97 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks72 = new A.ShapeLocks() { NoChangeAspect = true };

            nonVisualShapeDrawingProperties97.Append(shapeLocks72);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties138 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties97.Append(nonVisualDrawingProperties138);
            nonVisualShapeProperties97.Append(nonVisualShapeDrawingProperties97);
            nonVisualShapeProperties97.Append(applicationNonVisualDrawingProperties138);

            ShapeProperties shapeProperties116 = new ShapeProperties();

            A.Transform2D transform2D98 = new A.Transform2D();
            A.Offset offset120 = new A.Offset() { X = 9544250L, Y = 262305L };
            A.Extents extents120 = new A.Extents() { Cx = 1944000L, Cy = 1944000L };

            transform2D98.Append(offset120);
            transform2D98.Append(extents120);

            A.PresetGeometry presetGeometry64 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList64 = new A.AdjustValueList();

            presetGeometry64.Append(adjustValueList64);

            A.SolidFill solidFill129 = new A.SolidFill();

            A.SchemeColor schemeColor239 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.Alpha alpha16 = new A.Alpha() { Val = 70000 };

            schemeColor239.Append(alpha16);

            solidFill129.Append(schemeColor239);

            A.Outline outline35 = new A.Outline();
            A.NoFill noFill37 = new A.NoFill();

            outline35.Append(noFill37);
            A.EffectList effectList12 = new A.EffectList();

            shapeProperties116.Append(transform2D98);
            shapeProperties116.Append(presetGeometry64);
            shapeProperties116.Append(solidFill129);
            shapeProperties116.Append(outline35);
            shapeProperties116.Append(effectList12);

            ShapeStyle shapeStyle26 = new ShapeStyle();

            A.LineReference lineReference26 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor240 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade32 = new A.Shade() { Val = 50000 };

            schemeColor240.Append(shade32);

            lineReference26.Append(schemeColor240);

            A.FillReference fillReference26 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor241 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference26.Append(schemeColor241);

            A.EffectReference effectReference26 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor242 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference26.Append(schemeColor242);

            A.FontReference fontReference26 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor243 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference26.Append(schemeColor243);

            shapeStyle26.Append(lineReference26);
            shapeStyle26.Append(fillReference26);
            shapeStyle26.Append(effectReference26);
            shapeStyle26.Append(fontReference26);

            TextBody textBody91 = new TextBody();
            A.BodyProperties bodyProperties91 = new A.BodyProperties() { LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle91 = new A.ListStyle();

            A.Paragraph paragraph135 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties82 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter26 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints36 = new A.SpacingPoints() { Val = 400 };

            spaceAfter26.Append(spacingPoints36);

            paragraphProperties82.Append(spaceAfter26);

            A.Run run137 = new A.Run();

            A.RunProperties runProperties140 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill130 = new A.SolidFill();
            A.SchemeColor schemeColor244 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill130.Append(schemeColor244);
            A.LatinFont latinFont120 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.EastAsianFont eastAsianFont100 = new A.EastAsianFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont100 = new A.ComplexScriptFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties140.Append(solidFill130);
            runProperties140.Append(latinFont120);
            runProperties140.Append(eastAsianFont100);
            runProperties140.Append(complexScriptFont100);
            A.Text text139 = new A.Text();
            text139.Text = "#7Theme.Title";

            run137.Append(runProperties140);
            run137.Append(text139);

            paragraph135.Append(paragraphProperties82);
            paragraph135.Append(run137);

            A.Paragraph paragraph136 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties83 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter27 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints37 = new A.SpacingPoints() { Val = 400 };

            spaceAfter27.Append(spacingPoints37);

            paragraphProperties83.Append(spaceAfter27);

            A.EndParagraphRunProperties endParagraphRunProperties62 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill131 = new A.SolidFill();
            A.SchemeColor schemeColor245 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill131.Append(schemeColor245);
            A.LatinFont latinFont121 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties62.Append(solidFill131);
            endParagraphRunProperties62.Append(latinFont121);

            paragraph136.Append(paragraphProperties83);
            paragraph136.Append(endParagraphRunProperties62);

            A.Paragraph paragraph137 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties84 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter28 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints38 = new A.SpacingPoints() { Val = 400 };

            spaceAfter28.Append(spacingPoints38);

            paragraphProperties84.Append(spaceAfter28);

            A.Run run138 = new A.Run();

            A.RunProperties runProperties141 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill132 = new A.SolidFill();
            A.SchemeColor schemeColor246 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill132.Append(schemeColor246);
            A.LatinFont latinFont122 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties141.Append(solidFill132);
            runProperties141.Append(latinFont122);
            A.Text text140 = new A.Text();
            text140.Text = "#7Theme.Text";

            run138.Append(runProperties141);
            run138.Append(text140);

            paragraph137.Append(paragraphProperties84);
            paragraph137.Append(run138);

            A.Paragraph paragraph138 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties85 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter29 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints39 = new A.SpacingPoints() { Val = 400 };

            spaceAfter29.Append(spacingPoints39);

            paragraphProperties85.Append(spaceAfter29);

            A.EndParagraphRunProperties endParagraphRunProperties63 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill133 = new A.SolidFill();
            A.SchemeColor schemeColor247 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill133.Append(schemeColor247);
            A.LatinFont latinFont123 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties63.Append(solidFill133);
            endParagraphRunProperties63.Append(latinFont123);

            paragraph138.Append(paragraphProperties85);
            paragraph138.Append(endParagraphRunProperties63);

            A.Paragraph paragraph139 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties86 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter30 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints40 = new A.SpacingPoints() { Val = 400 };

            spaceAfter30.Append(spacingPoints40);

            paragraphProperties86.Append(spaceAfter30);

            A.Run run139 = new A.Run();

            A.RunProperties runProperties142 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill134 = new A.SolidFill();
            A.SchemeColor schemeColor248 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill134.Append(schemeColor248);
            A.LatinFont latinFont124 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties142.Append(solidFill134);
            runProperties142.Append(latinFont124);
            A.Text text141 = new A.Text();
            text141.Text = "#7Theme.SourceUrl";

            run139.Append(runProperties142);
            run139.Append(text141);

            A.EndParagraphRunProperties endParagraphRunProperties64 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 800, Dirty = false };

            A.SolidFill solidFill135 = new A.SolidFill();
            A.SchemeColor schemeColor249 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill135.Append(schemeColor249);

            endParagraphRunProperties64.Append(solidFill135);

            paragraph139.Append(paragraphProperties86);
            paragraph139.Append(run139);
            paragraph139.Append(endParagraphRunProperties64);

            textBody91.Append(bodyProperties91);
            textBody91.Append(listStyle91);
            textBody91.Append(paragraph135);
            textBody91.Append(paragraph136);
            textBody91.Append(paragraph137);
            textBody91.Append(paragraph138);
            textBody91.Append(paragraph139);

            shape97.Append(nonVisualShapeProperties97);
            shape97.Append(shapeProperties116);
            shape97.Append(shapeStyle26);
            shape97.Append(textBody91);
            */
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #8
            Shape shape98 = GenerateSpringboardDetailThemeOffset(27U, "Oval 26", 10008640L, 2326537L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[8].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[8].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[8].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #5
            Shape shape99 = GenerateSpringboardDetailThemeOffset(30U, "Oval 29", 8019715L, 1663300L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[5].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[5].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[5].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #3
            Shape shape100 = GenerateSpringboardDetailThemeOffset(31U, "Oval 30", 6027506L, 2548655L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[3].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[3].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[3].SourceUrl);
            #endregion

            #region Oval Shape, Theme Title, Text & SourceUrl #0
            Shape shape101 = GenerateSpringboardDetailThemeOffset(18U, "Oval 17", 4323901L, 1365983L,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[0].Title,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[0].Text,
                                         _springboard.Project.Areas[0].Springboards[0].Themes[0].SourceUrl);
            #endregion

            #region Big Picture & Url
            Shape shape102 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties102 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties143 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList8 = new A.NonVisualDrawingPropertiesExtensionList();
            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension8 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{34ED17CA-5441-4C65-B24E-CC1CE414383E}\" />");

            nonVisualDrawingPropertiesExtension8.Append(openXmlUnknownElement8);
            nonVisualDrawingPropertiesExtensionList8.Append(nonVisualDrawingPropertiesExtension8);
            nonVisualDrawingProperties143.Append(nonVisualDrawingPropertiesExtensionList8);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties102 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks77 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties102.Append(shapeLocks77);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties143 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape67 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties143.Append(placeholderShape67);

            nonVisualShapeProperties102.Append(nonVisualDrawingProperties143);
            nonVisualShapeProperties102.Append(nonVisualShapeDrawingProperties102);
            nonVisualShapeProperties102.Append(applicationNonVisualDrawingProperties143);
            ShapeProperties shapeProperties121 = new ShapeProperties();

            shape102.Append(nonVisualShapeProperties102);
            shape102.Append(shapeProperties121);
            #endregion

            // Oval Shape, Market #0
            Shape shape103 = GenerateSpringboardDetailMarketShapeOffset(36U, "Oval 35", 394740L, 6469080L, "FFF999");
            Shape shape104 = GenerateSpringboardDetailMarketTextOffset(4U, "TextBox 3", 632441L, 6513972L, _springboard.Project.Areas[0].Springboards[0].Themes[0].Market);
            // Oval Shape, Market #1
            Shape shape105 = GenerateSpringboardDetailMarketShapeOffset(46U, "Oval 45", 1425258L, 6469080L, "FFFF99");
            Shape shape106 = GenerateSpringboardDetailMarketTextOffset(47U, "TextBox 46", 1662959L, 6513972L, _springboard.Project.Areas[0].Springboards[0].Themes[1].Market);
            // Oval Shape, Market #2
            Shape shape107 = GenerateSpringboardDetailMarketShapeOffset(48U, "Oval 47", 2443103L, 6469080L, "99FF99");
            Shape shape108 = GenerateSpringboardDetailMarketTextOffset(49U, "TextBox 48", 2680804L, 6513972L, _springboard.Project.Areas[0].Springboards[0].Themes[2].Market);
            // Oval Shape, Market #3
            Shape shape109 = GenerateSpringboardDetailMarketShapeOffset(50U, "Oval 49", 3528415L, 6469080L, "99FF99");
            Shape shape110 = GenerateSpringboardDetailMarketTextOffset(51U, "TextBox 50", 3766116L, 6513972L, _springboard.Project.Areas[0].Springboards[0].Themes[3].Market);
            // Oval Shape, Market #4
            Shape shape111 = GenerateSpringboardDetailMarketShapeOffset(52U, "Oval 51", 4606495L, 6469080L, "FFCCFF");
            Shape shape112 = GenerateSpringboardDetailMarketTextOffset(53U, "TextBox 52", 4844196L, 6513972L, _springboard.Project.Areas[0].Springboards[0].Themes[4].Market);
            // Oval Shape, Market #5
            Shape shape113 = GenerateSpringboardDetailMarketShapeOffset(59U, "Oval 58", 8095693L, 6469080L, "CCCCFF");
            Shape shape114 = GenerateSpringboardDetailMarketTextOffset(60U, "TextBox 59", 8333394L, 6513972L, _springboard.Project.Areas[0].Springboards[0].Themes[5].Market);
            // Oval Shape, Market #6
            Shape shape115 = GenerateSpringboardDetailMarketShapeOffset(61U, "Oval 60", 9126211L, 6469080L, "CCCFFF");
            Shape shape116 = GenerateSpringboardDetailMarketTextOffset(62U, "TextBox 61", 9363912L, 6513972L, _springboard.Project.Areas[0].Springboards[0].Themes[6].Market);

            #region Project Teaser
            Shape shape117 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties117 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties158 = new NonVisualDrawingProperties() { Id = (UInt32Value)33U, Name = "TextBox 32" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList23 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension23 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement23 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{AD11E5F8-D73A-4B6D-83E4-7531C60882CF}\" />");

            nonVisualDrawingPropertiesExtension23.Append(openXmlUnknownElement23);

            nonVisualDrawingPropertiesExtensionList23.Append(nonVisualDrawingPropertiesExtension23);

            nonVisualDrawingProperties158.Append(nonVisualDrawingPropertiesExtensionList23);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties117 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties158 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties117.Append(nonVisualDrawingProperties158);
            nonVisualShapeProperties117.Append(nonVisualShapeDrawingProperties117);
            nonVisualShapeProperties117.Append(applicationNonVisualDrawingProperties158);

            ShapeProperties shapeProperties136 = new ShapeProperties();

            A.Transform2D transform2D117 = new A.Transform2D();
            A.Offset offset139 = new A.Offset() { X = 621708L, Y = 4121491L };
            A.Extents extents139 = new A.Extents() { Cx = 1251240L, Cy = 276999L };

            transform2D117.Append(offset139);
            transform2D117.Append(extents139);

            A.PresetGeometry presetGeometry83 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList83 = new A.AdjustValueList();

            presetGeometry83.Append(adjustValueList83);
            A.NoFill noFill56 = new A.NoFill();

            shapeProperties136.Append(transform2D117);
            shapeProperties136.Append(presetGeometry83);
            shapeProperties136.Append(noFill56);

            TextBody textBody110 = new TextBody();

            A.BodyProperties bodyProperties110 = new A.BodyProperties() { Wrap = A.TextWrappingValues.None, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit17 = new A.ShapeAutoFit();

            bodyProperties110.Append(shapeAutoFit17);
            A.ListStyle listStyle110 = new A.ListStyle();

            A.Paragraph paragraph174 = new A.Paragraph();

            A.Run run159 = new A.Run();

            A.RunProperties runProperties162 = new A.RunProperties() { Language = "en-GB", FontSize = 1200, Dirty = false };

            A.SolidFill solidFill185 = new A.SolidFill();
            A.SchemeColor schemeColor339 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill185.Append(schemeColor339);
            A.LatinFont latinFont159 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", Panose = "020F0704030504030204", PitchFamily = 34, CharacterSet = 0 };

            A.Run run160 = new A.Run();
            A.RunProperties runProperties163 = new A.RunProperties() { Language = "en-GB", FontSize = 1200, Dirty = false, SpellingError = true };

            A.SolidFill solidFill186 = new A.SolidFill();
            A.SchemeColor schemeColor340 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill186.Append(schemeColor340);
            A.LatinFont latinFont160 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", Panose = "020F0704030504030204", PitchFamily = 34, CharacterSet = 0 };

            runProperties163.Append(solidFill186);
            runProperties163.Append(latinFont160);
            A.Text text162 = new A.Text();
            text162.Text = _springboard.Project.Teaser;//"Project.Teaser";

            run160.Append(runProperties163);
            run160.Append(text162);

            A.EndParagraphRunProperties endParagraphRunProperties84 = new A.EndParagraphRunProperties() { Language = "en-GB", FontSize = 1200, Dirty = false };

            A.SolidFill solidFill187 = new A.SolidFill();
            A.SchemeColor schemeColor341 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill187.Append(schemeColor341);
            A.LatinFont latinFont161 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", Panose = "020F0704030504030204", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties84.Append(solidFill187);
            endParagraphRunProperties84.Append(latinFont161);

            paragraph174.Append(run159);
            paragraph174.Append(run160);
            paragraph174.Append(endParagraphRunProperties84);

            textBody110.Append(bodyProperties110);
            textBody110.Append(listStyle110);
            textBody110.Append(paragraph174);

            shape117.Append(nonVisualShapeProperties117);
            shape117.Append(shapeProperties136);
            shape117.Append(textBody110);
            #endregion

            #region Springboard Picture Url
            Shape shape118 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties118 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties159 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "TextBox 4" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties118 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties159 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties118.Append(nonVisualDrawingProperties159);
            nonVisualShapeProperties118.Append(nonVisualShapeDrawingProperties118);
            nonVisualShapeProperties118.Append(applicationNonVisualDrawingProperties159);

            ShapeProperties shapeProperties137 = new ShapeProperties();

            A.Transform2D transform2D118 = new A.Transform2D();
            A.Offset offset140 = new A.Offset() { X = 962820L, Y = 1365983L };
            A.Extents extents140 = new A.Extents() { Cx = 1785512L, Cy = 646331L };

            transform2D118.Append(offset140);
            transform2D118.Append(extents140);

            A.PresetGeometry presetGeometry84 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList84 = new A.AdjustValueList();

            presetGeometry84.Append(adjustValueList84);
            A.NoFill noFill57 = new A.NoFill();

            shapeProperties137.Append(transform2D118);
            shapeProperties137.Append(presetGeometry84);
            shapeProperties137.Append(noFill57);

            TextBody textBody111 = new TextBody();

            A.BodyProperties bodyProperties111 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit18 = new A.ShapeAutoFit();

            bodyProperties111.Append(shapeAutoFit18);
            A.ListStyle listStyle111 = new A.ListStyle();

            A.Paragraph paragraph175 = new A.Paragraph();

            A.Run run161 = new A.Run();
            A.RunProperties runProperties164 = new A.RunProperties() { Language = "en-GB" };
            A.Text text163 = new A.Text();
            text163.Text = _springboard.Project.Areas[0].Springboards[0].ImageUrl; //"$Springboard.PictureUrl";

            run161.Append(runProperties164);
            run161.Append(text163);
            A.EndParagraphRunProperties endParagraphRunProperties85 = new A.EndParagraphRunProperties() { Language = "en-GB", Dirty = false };

            paragraph175.Append(run161);
            paragraph175.Append(endParagraphRunProperties85);

            textBody111.Append(bodyProperties111);
            textBody111.Append(listStyle111);
            textBody111.Append(paragraph175);

            shape118.Append(nonVisualShapeProperties118);
            shape118.Append(shapeProperties137);
            shape118.Append(textBody111);
            #endregion

            shapeTree17.Append(nonVisualGroupShapeProperties22);
            shapeTree17.Append(groupShapeProperties22);
            shapeTree17.Append(shape89);
            shapeTree17.Append(shape90);
            shapeTree17.Append(shape91);
            shapeTree17.Append(picture15);
            shapeTree17.Append(shape92);
            shapeTree17.Append(shape93);
            shapeTree17.Append(shape94);
            shapeTree17.Append(shape95);
            shapeTree17.Append(shape96);
            shapeTree17.Append(shape97);
            shapeTree17.Append(shape98);
            shapeTree17.Append(shape99);
            shapeTree17.Append(shape100);
            shapeTree17.Append(shape101);
            shapeTree17.Append(shape102);
            shapeTree17.Append(shape103);
            shapeTree17.Append(shape104);
            shapeTree17.Append(shape105);
            shapeTree17.Append(shape106);
            shapeTree17.Append(shape107);
            shapeTree17.Append(shape108);
            shapeTree17.Append(shape109);
            shapeTree17.Append(shape110);
            shapeTree17.Append(shape111);
            shapeTree17.Append(shape112);
            shapeTree17.Append(shape113);
            shapeTree17.Append(shape114);
            shapeTree17.Append(shape115);
            shapeTree17.Append(shape116);
            shapeTree17.Append(shape117);
            shapeTree17.Append(shape118);

            CommonSlideDataExtensionList commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension11 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId11 = new P14.CreationId() { Val = (UInt32Value)1356576464U };
            creationId11.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension11.Append(creationId11);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension11);

            commonSlideData17.Append(shapeTree17);
            commonSlideData17.Append(commonSlideDataExtensionList12);

            ColorMapOverride colorMapOverride15 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping15 = new A.MasterColorMapping();

            colorMapOverride15.Append(masterColorMapping15);

            slide6.Append(commonSlideData17);
            slide6.Append(colorMapOverride15);

            slidePart6.Slide = slide6;
        }
        #endregion

        private Shape GenerateSpringboardDetailMarketTextOffset(UInt32Value idText, string nameText, long xOffsetText, long yOffsetText, string marketText)
        {
            Shape shape104 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties104 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties145 = new NonVisualDrawingProperties() { Id = idText, Name = nameText };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList10 = new A.NonVisualDrawingPropertiesExtensionList();
            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension10 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement10 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9E4D0FEA-4BD6-4AB1-B215-6F833CB0A807}\" />");

            nonVisualDrawingPropertiesExtension10.Append(openXmlUnknownElement10);
            nonVisualDrawingPropertiesExtensionList10.Append(nonVisualDrawingPropertiesExtension10);

            nonVisualDrawingProperties145.Append(nonVisualDrawingPropertiesExtensionList10);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties104 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties145 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties104.Append(nonVisualDrawingProperties145);
            nonVisualShapeProperties104.Append(nonVisualShapeDrawingProperties104);
            nonVisualShapeProperties104.Append(applicationNonVisualDrawingProperties145);

            ShapeProperties shapeProperties123 = new ShapeProperties();

            A.Transform2D transform2D104 = new A.Transform2D();
            A.Offset offset126 = new A.Offset() { X = xOffsetText, Y = yOffsetText };
            A.Extents extents126 = new A.Extents() { Cx = 614271L, Cy = 215444L };

            transform2D104.Append(offset126);
            transform2D104.Append(extents126);

            A.PresetGeometry presetGeometry70 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList70 = new A.AdjustValueList();

            presetGeometry70.Append(adjustValueList70);
            A.NoFill noFill43 = new A.NoFill();

            shapeProperties123.Append(transform2D104);
            shapeProperties123.Append(presetGeometry70);
            shapeProperties123.Append(noFill43);

            TextBody textBody97 = new TextBody();

            A.BodyProperties bodyProperties97 = new A.BodyProperties() { Wrap = A.TextWrappingValues.None, RightToLeftColumns = false };
            A.ShapeAutoFit shapeAutoFit10 = new A.ShapeAutoFit();

            bodyProperties97.Append(shapeAutoFit10);
            A.ListStyle listStyle97 = new A.ListStyle();

            A.Paragraph paragraph161 = new A.Paragraph();

            A.Run run152 = new A.Run();

            A.RunProperties runProperties155 = new A.RunProperties() { Language = "en-GB", FontSize = 800, Dirty = false };

            A.SolidFill solidFill166 = new A.SolidFill();
            A.SchemeColor schemeColor300 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill166.Append(schemeColor300);
            A.LatinFont latinFont146 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", Panose = "020F0704030504030204", PitchFamily = 34, CharacterSet = 0 };

            runProperties155.Append(solidFill166);
            runProperties155.Append(latinFont146);
            A.Text text154 = new A.Text();
            text154.Text = marketText; //"#0Market";

            run152.Append(runProperties155);
            run152.Append(text154);

            paragraph161.Append(run152);

            textBody97.Append(bodyProperties97);
            textBody97.Append(listStyle97);
            textBody97.Append(paragraph161);

            shape104.Append(nonVisualShapeProperties104);
            shape104.Append(shapeProperties123);
            shape104.Append(textBody97);

            return shape104;
        }

        private Shape GenerateSpringboardDetailMarketShapeOffset(UInt32Value idOval, string nameOval, long xOffsetOval, long yOffsetOval, string colourOval)
        {
            Shape shape105 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties105 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties146 = new NonVisualDrawingProperties() { Id = idOval, Name = nameOval };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList11 = new A.NonVisualDrawingPropertiesExtensionList();
            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension11 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement11 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2BFC2031-CD86-4DAD-84E3-12279D47509F}\" />");

            nonVisualDrawingPropertiesExtension11.Append(openXmlUnknownElement11);
            nonVisualDrawingPropertiesExtensionList11.Append(nonVisualDrawingPropertiesExtension11);

            nonVisualDrawingProperties146.Append(nonVisualDrawingPropertiesExtensionList11);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties105 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties146 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties105.Append(nonVisualDrawingProperties146);
            nonVisualShapeProperties105.Append(nonVisualShapeDrawingProperties105);
            nonVisualShapeProperties105.Append(applicationNonVisualDrawingProperties146);

            ShapeProperties shapeProperties124 = new ShapeProperties();

            A.Transform2D transform2D105 = new A.Transform2D();
            A.Offset offset127 = new A.Offset() { X = xOffsetOval, Y = yOffsetOval };
            A.Extents extents127 = new A.Extents() { Cx = 305229L, Cy = 305229L };

            transform2D105.Append(offset127);
            transform2D105.Append(extents127);

            A.PresetGeometry presetGeometry71 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList71 = new A.AdjustValueList();

            presetGeometry71.Append(adjustValueList71);

            A.SolidFill solidFill167 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex43 = new A.RgbColorModelHex() { Val = colourOval };

            solidFill167.Append(rgbColorModelHex43);

            A.Outline outline41 = new A.Outline();
            A.NoFill noFill44 = new A.NoFill();

            outline41.Append(noFill44);
            A.EffectList effectList18 = new A.EffectList();

            shapeProperties124.Append(transform2D105);
            shapeProperties124.Append(presetGeometry71);
            shapeProperties124.Append(solidFill167);
            shapeProperties124.Append(outline41);
            shapeProperties124.Append(effectList18);

            ShapeStyle shapeStyle32 = new ShapeStyle();

            A.LineReference lineReference32 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor301 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade38 = new A.Shade() { Val = 50000 };

            schemeColor301.Append(shade38);

            lineReference32.Append(schemeColor301);

            A.FillReference fillReference32 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor302 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference32.Append(schemeColor302);

            A.EffectReference effectReference32 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor303 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference32.Append(schemeColor303);

            A.FontReference fontReference32 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor304 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference32.Append(schemeColor304);

            shapeStyle32.Append(lineReference32);
            shapeStyle32.Append(fillReference32);
            shapeStyle32.Append(effectReference32);
            shapeStyle32.Append(fontReference32);

            TextBody textBody98 = new TextBody();
            A.BodyProperties bodyProperties98 = new A.BodyProperties() { LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle98 = new A.ListStyle();

            A.Paragraph paragraph162 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties108 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter52 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints62 = new A.SpacingPoints() { Val = 400 };

            spaceAfter52.Append(spacingPoints62);

            paragraphProperties108.Append(spaceAfter52);

            A.EndParagraphRunProperties endParagraphRunProperties78 = new A.EndParagraphRunProperties() { Language = "en-GB", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill168 = new A.SolidFill();
            A.SchemeColor schemeColor305 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill168.Append(schemeColor305);
            A.LatinFont latinFont147 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties78.Append(solidFill168);
            endParagraphRunProperties78.Append(latinFont147);

            paragraph162.Append(paragraphProperties108);
            paragraph162.Append(endParagraphRunProperties78);

            textBody98.Append(bodyProperties98);
            textBody98.Append(listStyle98);
            textBody98.Append(paragraph162);

            shape105.Append(nonVisualShapeProperties105);
            shape105.Append(shapeProperties124);
            shape105.Append(shapeStyle32);
            shape105.Append(textBody98);

            return shape105;
        }

        private Shape GenerateSpringboardDetailThemeOffset(UInt32Value shapeId, string shapeName, long xOffset, long yOffset, string themeTitle, string themeText, string themeSourceUrl)
        {
            Shape shape92 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties92 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties133 = new NonVisualDrawingProperties() { Id = shapeId, Name = shapeName };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties92 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks67 = new A.ShapeLocks() { NoChangeAspect = true };

            nonVisualShapeDrawingProperties92.Append(shapeLocks67);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties133 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties92.Append(nonVisualDrawingProperties133);
            nonVisualShapeProperties92.Append(nonVisualShapeDrawingProperties92);
            nonVisualShapeProperties92.Append(applicationNonVisualDrawingProperties133);

            ShapeProperties shapeProperties111 = new ShapeProperties();

            A.Transform2D transform2D93 = new A.Transform2D();
            A.Offset offset115 = new A.Offset() { X = xOffset, Y = yOffset };
            A.Extents extents115 = new A.Extents() { Cx = 1944000L, Cy = 1944000L };

            transform2D93.Append(offset115);
            transform2D93.Append(extents115);

            A.PresetGeometry presetGeometry59 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList59 = new A.AdjustValueList();

            presetGeometry59.Append(adjustValueList59);

            A.SolidFill solidFill94 = new A.SolidFill();

            A.SchemeColor schemeColor184 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };
            A.Alpha alpha11 = new A.Alpha() { Val = 70000 };

            schemeColor184.Append(alpha11);

            solidFill94.Append(schemeColor184);

            A.Outline outline30 = new A.Outline();
            A.NoFill noFill32 = new A.NoFill();

            outline30.Append(noFill32);
            A.EffectList effectList7 = new A.EffectList();

            shapeProperties111.Append(transform2D93);
            shapeProperties111.Append(presetGeometry59);
            shapeProperties111.Append(solidFill94);
            shapeProperties111.Append(outline30);
            shapeProperties111.Append(effectList7);

            ShapeStyle shapeStyle21 = new ShapeStyle();

            A.LineReference lineReference21 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor185 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade27 = new A.Shade() { Val = 50000 };

            schemeColor185.Append(shade27);

            lineReference21.Append(schemeColor185);

            A.FillReference fillReference21 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor186 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference21.Append(schemeColor186);

            A.EffectReference effectReference21 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor187 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference21.Append(schemeColor187);

            A.FontReference fontReference21 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor188 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference21.Append(schemeColor188);

            shapeStyle21.Append(lineReference21);
            shapeStyle21.Append(fillReference21);
            shapeStyle21.Append(effectReference21);
            shapeStyle21.Append(fontReference21);

            TextBody textBody86 = new TextBody();
            A.BodyProperties bodyProperties86 = new A.BodyProperties() { LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle86 = new A.ListStyle();

            A.Paragraph paragraph110 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties57 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter1 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints11 = new A.SpacingPoints() { Val = 400 };

            spaceAfter1.Append(spacingPoints11);

            paragraphProperties57.Append(spaceAfter1);

            A.Run run122 = new A.Run();

            A.RunProperties runProperties125 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill95 = new A.SolidFill();
            A.SchemeColor schemeColor189 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill95.Append(schemeColor189);
            A.LatinFont latinFont95 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties125.Append(solidFill95);
            runProperties125.Append(latinFont95);
            A.Text text124 = new A.Text();
            text124.Text = themeTitle; // "#6Theme.Title";

            run122.Append(runProperties125);
            run122.Append(text124);

            paragraph110.Append(paragraphProperties57);
            paragraph110.Append(run122);

            A.Paragraph paragraph111 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties58 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter2 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints12 = new A.SpacingPoints() { Val = 400 };

            spaceAfter2.Append(spacingPoints12);

            paragraphProperties58.Append(spaceAfter2);

            A.EndParagraphRunProperties endParagraphRunProperties47 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill96 = new A.SolidFill();
            A.SchemeColor schemeColor190 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill96.Append(schemeColor190);
            A.LatinFont latinFont96 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties47.Append(solidFill96);
            endParagraphRunProperties47.Append(latinFont96);

            paragraph111.Append(paragraphProperties58);
            paragraph111.Append(endParagraphRunProperties47);

            A.Paragraph paragraph112 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties59 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter3 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints13 = new A.SpacingPoints() { Val = 400 };

            spaceAfter3.Append(spacingPoints13);

            paragraphProperties59.Append(spaceAfter3);

            A.Run run123 = new A.Run();

            A.RunProperties runProperties126 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill97 = new A.SolidFill();
            A.SchemeColor schemeColor191 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill97.Append(schemeColor191);
            A.LatinFont latinFont97 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties126.Append(solidFill97);
            runProperties126.Append(latinFont97);
            A.Text text125 = new A.Text();
            text125.Text = themeText; //"#6Theme.Text";

            run123.Append(runProperties126);
            run123.Append(text125);

            paragraph112.Append(paragraphProperties59);
            paragraph112.Append(run123);

            A.Paragraph paragraph113 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties60 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter4 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints14 = new A.SpacingPoints() { Val = 400 };

            spaceAfter4.Append(spacingPoints14);

            paragraphProperties60.Append(spaceAfter4);

            A.EndParagraphRunProperties endParagraphRunProperties48 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill98 = new A.SolidFill();
            A.SchemeColor schemeColor192 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill98.Append(schemeColor192);
            A.LatinFont latinFont98 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            endParagraphRunProperties48.Append(solidFill98);
            endParagraphRunProperties48.Append(latinFont98);

            paragraph113.Append(paragraphProperties60);
            paragraph113.Append(endParagraphRunProperties48);

            A.Paragraph paragraph114 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties61 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.SpaceAfter spaceAfter5 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints15 = new A.SpacingPoints() { Val = 400 };

            spaceAfter5.Append(spacingPoints15);

            paragraphProperties61.Append(spaceAfter5);

            A.Run run124 = new A.Run();

            A.RunProperties runProperties127 = new A.RunProperties() { Language = "en-US", FontSize = 1100, Dirty = false };

            A.SolidFill solidFill99 = new A.SolidFill();
            A.SchemeColor schemeColor193 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill99.Append(schemeColor193);
            A.LatinFont latinFont99 = new A.LatinFont() { Typeface = "Arial Rounded MT Bold", CharacterSet = 0 };

            runProperties127.Append(solidFill99);
            runProperties127.Append(latinFont99);
            A.Text text126 = new A.Text();
            text126.Text = themeSourceUrl; //"#6Theme.SourceUrl";

            run124.Append(runProperties127);
            run124.Append(text126);

            A.EndParagraphRunProperties endParagraphRunProperties49 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 800, Dirty = false };

            A.SolidFill solidFill100 = new A.SolidFill();
            A.SchemeColor schemeColor194 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill100.Append(schemeColor194);

            endParagraphRunProperties49.Append(solidFill100);

            paragraph114.Append(paragraphProperties61);
            paragraph114.Append(run124);
            paragraph114.Append(endParagraphRunProperties49);

            textBody86.Append(bodyProperties86);
            textBody86.Append(listStyle86);
            textBody86.Append(paragraph110);
            textBody86.Append(paragraph111);
            textBody86.Append(paragraph112);
            textBody86.Append(paragraph113);
            textBody86.Append(paragraph114);

            shape92.Append(nonVisualShapeProperties92);
            shape92.Append(shapeProperties111);
            shape92.Append(shapeStyle21);
            shape92.Append(textBody86);

            return shape92;
        }

        // Generates content of notesSlidePart1.
        private void GenerateNotesSlidePart1Content(NotesSlidePart notesSlidePart1)
        {
            NotesSlide notesSlide1 = new NotesSlide();
            notesSlide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            notesSlide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            notesSlide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData18 = new CommonSlideData();

            ShapeTree shapeTree18 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties23 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties160 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties23 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties160 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties23.Append(nonVisualDrawingProperties160);
            nonVisualGroupShapeProperties23.Append(nonVisualGroupShapeDrawingProperties23);
            nonVisualGroupShapeProperties23.Append(applicationNonVisualDrawingProperties160);

            GroupShapeProperties groupShapeProperties23 = new GroupShapeProperties();

            A.TransformGroup transformGroup23 = new A.TransformGroup();
            A.Offset offset141 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents141 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset23 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents23 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup23.Append(offset141);
            transformGroup23.Append(extents141);
            transformGroup23.Append(childOffset23);
            transformGroup23.Append(childExtents23);

            groupShapeProperties23.Append(transformGroup23);

            Shape shape119 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties119 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties161 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Slide Image Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties119 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks78 = new A.ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true };

            nonVisualShapeDrawingProperties119.Append(shapeLocks78);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties161 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape68 = new PlaceholderShape() { Type = PlaceholderValues.SlideImage };

            applicationNonVisualDrawingProperties161.Append(placeholderShape68);

            nonVisualShapeProperties119.Append(nonVisualDrawingProperties161);
            nonVisualShapeProperties119.Append(nonVisualShapeDrawingProperties119);
            nonVisualShapeProperties119.Append(applicationNonVisualDrawingProperties161);
            ShapeProperties shapeProperties138 = new ShapeProperties();

            shape119.Append(nonVisualShapeProperties119);
            shape119.Append(shapeProperties138);

            Shape shape120 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties120 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties162 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Notes Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties120 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks79 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties120.Append(shapeLocks79);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties162 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape69 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties162.Append(placeholderShape69);

            nonVisualShapeProperties120.Append(nonVisualDrawingProperties162);
            nonVisualShapeProperties120.Append(nonVisualShapeDrawingProperties120);
            nonVisualShapeProperties120.Append(applicationNonVisualDrawingProperties162);
            ShapeProperties shapeProperties139 = new ShapeProperties();

            TextBody textBody112 = new TextBody();
            A.BodyProperties bodyProperties112 = new A.BodyProperties();
            A.ListStyle listStyle112 = new A.ListStyle();

            A.Paragraph paragraph176 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties86 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph176.Append(endParagraphRunProperties86);

            textBody112.Append(bodyProperties112);
            textBody112.Append(listStyle112);
            textBody112.Append(paragraph176);

            shape120.Append(nonVisualShapeProperties120);
            shape120.Append(shapeProperties139);
            shape120.Append(textBody112);

            Shape shape121 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties121 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties163 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Slide Number Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties121 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks80 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties121.Append(shapeLocks80);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties163 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape70 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties163.Append(placeholderShape70);

            nonVisualShapeProperties121.Append(nonVisualDrawingProperties163);
            nonVisualShapeProperties121.Append(nonVisualShapeDrawingProperties121);
            nonVisualShapeProperties121.Append(applicationNonVisualDrawingProperties163);
            ShapeProperties shapeProperties140 = new ShapeProperties();

            TextBody textBody113 = new TextBody();
            A.BodyProperties bodyProperties113 = new A.BodyProperties();
            A.ListStyle listStyle113 = new A.ListStyle();

            A.Paragraph paragraph177 = new A.Paragraph();

            A.Field field3 = new A.Field() { Id = "{D7BFF79C-DF71-994D-85D5-70B4949C3937}", Type = "slidenum" };

            A.RunProperties runProperties165 = new A.RunProperties() { Language = "en-US" };
            runProperties165.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text164 = new A.Text();
            text164.Text = "‹#›";

            field3.Append(runProperties165);
            field3.Append(text164);
            A.EndParagraphRunProperties endParagraphRunProperties87 = new A.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph177.Append(field3);
            paragraph177.Append(endParagraphRunProperties87);

            textBody113.Append(bodyProperties113);
            textBody113.Append(listStyle113);
            textBody113.Append(paragraph177);

            shape121.Append(nonVisualShapeProperties121);
            shape121.Append(shapeProperties140);
            shape121.Append(textBody113);

            shapeTree18.Append(nonVisualGroupShapeProperties23);
            shapeTree18.Append(groupShapeProperties23);
            shapeTree18.Append(shape119);
            shapeTree18.Append(shape120);
            shapeTree18.Append(shape121);

            CommonSlideDataExtensionList commonSlideDataExtensionList13 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension12 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId12 = new P14.CreationId() { Val = (UInt32Value)2500277414U };
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList13.Append(commonSlideDataExtension12);

            commonSlideData18.Append(shapeTree18);
            commonSlideData18.Append(commonSlideDataExtensionList13);

            ColorMapOverride colorMapOverride16 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping16 = new A.MasterColorMapping();

            colorMapOverride16.Append(masterColorMapping16);

            notesSlide1.Append(commonSlideData18);
            notesSlide1.Append(colorMapOverride16);

            notesSlidePart1.NotesSlide = notesSlide1;
        }

        // Generates content of presentationPropertiesPart1.
        private void GeneratePresentationPropertiesPart1Content(PresentationPropertiesPart presentationPropertiesPart1)
        {
            PresentationProperties presentationProperties1 = new PresentationProperties();
            presentationProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentationProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentationProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            ShowProperties showProperties1 = new ShowProperties() { ShowNarration = true };
            PresenterSlideMode presenterSlideMode1 = new PresenterSlideMode();
            SlideAll slideAll1 = new SlideAll();

            PenColor penColor1 = new PenColor();
            A.PresetColor presetColor2 = new A.PresetColor() { Val = A.PresetColorValues.Red };

            penColor1.Append(presetColor2);

            ShowPropertiesExtensionList showPropertiesExtensionList1 = new ShowPropertiesExtensionList();

            ShowPropertiesExtension showPropertiesExtension1 = new ShowPropertiesExtension() { Uri = "{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}" };

            P14.LaserColor laserColor1 = new P14.LaserColor();
            laserColor1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
            A.RgbColorModelHex rgbColorModelHex47 = new A.RgbColorModelHex() { Val = "FF0000" };

            laserColor1.Append(rgbColorModelHex47);

            showPropertiesExtension1.Append(laserColor1);

            ShowPropertiesExtension showPropertiesExtension2 = new ShowPropertiesExtension() { Uri = "{2FDB2607-1784-4EEB-B798-7EB5836EED8A}" };

            P14.ShowMediaControls showMediaControls1 = new P14.ShowMediaControls() { Val = true };
            showMediaControls1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            showPropertiesExtension2.Append(showMediaControls1);

            showPropertiesExtensionList1.Append(showPropertiesExtension1);
            showPropertiesExtensionList1.Append(showPropertiesExtension2);

            showProperties1.Append(presenterSlideMode1);
            showProperties1.Append(slideAll1);
            showProperties1.Append(penColor1);
            showProperties1.Append(showPropertiesExtensionList1);

            ColorMostRecentlyUsed colorMostRecentlyUsed1 = new ColorMostRecentlyUsed();
            A.RgbColorModelHex rgbColorModelHex48 = new A.RgbColorModelHex() { Val = "86D5AC" };
            A.RgbColorModelHex rgbColorModelHex49 = new A.RgbColorModelHex() { Val = "6CA684" };

            colorMostRecentlyUsed1.Append(rgbColorModelHex48);
            colorMostRecentlyUsed1.Append(rgbColorModelHex49);

            PresentationPropertiesExtensionList presentationPropertiesExtensionList1 = new PresentationPropertiesExtensionList();

            PresentationPropertiesExtension presentationPropertiesExtension1 = new PresentationPropertiesExtension() { Uri = "{E76CE94A-603C-4142-B9EB-6D1370010A27}" };

            P14.DiscardImageEditData discardImageEditData1 = new P14.DiscardImageEditData() { Val = false };
            discardImageEditData1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension1.Append(discardImageEditData1);

            PresentationPropertiesExtension presentationPropertiesExtension2 = new PresentationPropertiesExtension() { Uri = "{D31A062A-798A-4329-ABDD-BBA856620510}" };

            P14.DefaultImageDpi defaultImageDpi1 = new P14.DefaultImageDpi() { Val = (UInt32Value)32767U };
            defaultImageDpi1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension2.Append(defaultImageDpi1);

            PresentationPropertiesExtension presentationPropertiesExtension3 = new PresentationPropertiesExtension() { Uri = "{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}" };

            P15.ChartTrackingReferenceBased chartTrackingReferenceBased1 = new P15.ChartTrackingReferenceBased() { Val = false };
            chartTrackingReferenceBased1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            presentationPropertiesExtension3.Append(chartTrackingReferenceBased1);

            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension1);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension2);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension3);

            presentationProperties1.Append(showProperties1);
            presentationProperties1.Append(colorMostRecentlyUsed1);
            presentationProperties1.Append(presentationPropertiesExtensionList1);

            presentationPropertiesPart1.PresentationProperties = presentationProperties1;
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Office Theme";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "811";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "319";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office PowerPoint";
            Ap.PresentationFormat presentationFormat1 = new Ap.PresentationFormat();
            presentationFormat1.Text = "Widescreen";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "113";
            Ap.Slides slides1 = new Ap.Slides();
            slides1.Text = "6";
            Ap.Notes notes1 = new Ap.Notes();
            notes1.Text = "1";
            Ap.HiddenSlides hiddenSlides1 = new Ap.HiddenSlides();
            hiddenSlides1.Text = "0";
            Ap.MultimediaClips multimediaClips1 = new Ap.MultimediaClips();
            multimediaClips1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)6U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Fonts Used";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "4";

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Theme";

            variant3.Append(vTLPSTR2);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "1";

            variant4.Append(vTInt322);

            Vt.Variant variant5 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Slide Titles";

            variant5.Append(vTLPSTR3);

            Vt.Variant variant6 = new Vt.Variant();
            Vt.VTInt32 vTInt323 = new Vt.VTInt32();
            vTInt323.Text = "6";

            variant6.Append(vTInt323);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);
            vTVector1.Append(variant5);
            vTVector1.Append(variant6);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)11U };
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = ".AppleSystemUIFont";
            Vt.VTLPSTR vTLPSTR5 = new Vt.VTLPSTR();
            vTLPSTR5.Text = "Arial";
            Vt.VTLPSTR vTLPSTR6 = new Vt.VTLPSTR();
            vTLPSTR6.Text = "Arial Rounded MT Bold";
            Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR();
            vTLPSTR7.Text = "Calibri";
            Vt.VTLPSTR vTLPSTR8 = new Vt.VTLPSTR();
            vTLPSTR8.Text = "Office Theme";
            Vt.VTLPSTR vTLPSTR9 = new Vt.VTLPSTR();
            vTLPSTR9.Text = "$Question";
            Vt.VTLPSTR vTLPSTR10 = new Vt.VTLPSTR();
            vTLPSTR10.Text = "$Area.TitleSpringboards";
            Vt.VTLPSTR vTLPSTR11 = new Vt.VTLPSTR();
            vTLPSTR11.Text = "$SpringBoard.Title";
            Vt.VTLPSTR vTLPSTR12 = new Vt.VTLPSTR();
            vTLPSTR12.Text = "$WordCloud.Title";
            Vt.VTLPSTR vTLPSTR13 = new Vt.VTLPSTR();
            vTLPSTR13.Text = "$WordList.Title";
            Vt.VTLPSTR vTLPSTR14 = new Vt.VTLPSTR();
            vTLPSTR14.Text = "Project Sources";

            vTVector2.Append(vTLPSTR4);
            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);
            vTVector2.Append(vTLPSTR7);
            vTVector2.Append(vTLPSTR8);
            vTVector2.Append(vTLPSTR9);
            vTVector2.Append(vTLPSTR10);
            vTVector2.Append(vTLPSTR11);
            vTVector2.Append(vTLPSTR12);
            vTVector2.Append(vTLPSTR13);
            vTVector2.Append(vTLPSTR14);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(words1);
            properties1.Append(application1);
            properties1.Append(presentationFormat1);
            properties1.Append(paragraphs1);
            properties1.Append(slides1);
            properties1.Append(notes1);
            properties1.Append(hiddenSlides1);
            properties1.Append(multimediaClips1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Microsoft Office User";
            document.PackageProperties.Title = "PowerPoint Presentation";
            document.PackageProperties.Revision = "101";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2018-02-19T09:20:06Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2018-03-15T21:30:12Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "User";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2018-02-20T09:33:08Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #region Binary Data
        private string thumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACQAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD82vHX7QXjbwX8VfjdZQfELUdRnYMtpNpc1vdWWiWKapDBcSvbNGJba4tbaV/mIaNyrOxOUA7rw5+01b/AA0+M0/h7wt8QW+IXw1kvtHstS1vxLqJuxot1cSX8cyLdnbkOLWF13lkG7K/K4r7VtfCWiWWtanq8Gk2cOqanGkV7eJAokuUQEIsjYywAYjn1rPsPhh4O0q4lnsvCeh2c00Qgkkt9OhjZ4w/mBGIXld/zYPGeetC6A+p8Z+Ev2+/HnjDWtJSz8PeF1spY4mkjvL+Gz+2pLNeILi3lmu1fZGtspKxwTlyJcFCFB5uf9tb4i+MtN8NOdc0DQp4dRvINXjtLIJBOr6Hc3VvFHcR3twjEyROEeOUMzmBmjXmN/0Ih0PTbea3mi0+1jltwwhkSFQ0YY5YKccZJ5x1qL/hGdH+yfZf7Jsfsvm+f5P2ZNnmf39uMbvfrQM+ILX9uTxlo/g0aitv4dvZrS1lsx4buPPfWd0OgnURqcr+b81s7qFI8sfK4bzSx219Efs9/FvxT8Qte8Y6N4ri0f7To8Wl3lvPo8EsKGK9tBOInWSRyzRncu8FQ4wdi9K9bOg6YWYnTrQlrf7IT5C8w/8APLp9z/Z6VPb6fa2ckkkFtDDJIFV2jjClgowoJHXA4HpTJLFFFFIYUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFctD8UPCk2pvp/9uWsN4jFGiuGMJ3Db8uXAG7LAY65yOqnE3/Cx/CqzvA/iTSoZ0llgaGa8jjkDxPskXaxB+VsA8dx6igDo6K5m5+JnhGzura2m8T6QlxcDfFH9tjJZdu7d14XH8R45AzkjLl+JXhRvI3eI9Mi8+1jvovOukj3wOCyyDcRlSFJ9gOaAOkorlJPiv4MjuLKA+KdIaS9Vng2XkbK6hQxbcDgDBBBJGcjHWrbfELwutnPdf8ACR6S1tAFaWRb2NlQMzKmcN/EysB6lSByKAOgornB8RvC7NqKx6/p8z6faNf3SwzrI0cC53SYXJKjBBx0OB3FPk+IHhq3gtZrjXbCzS6tUvYftk6wF4WBKuA5Bxwfp3xQB0FFc23xK8JhlUeJtJcsGKhL2Ns4GSBg9cc468H0NOt/iL4WvYVltfEel3sbKrqbW7jmypkWIN8pPy73Vd3QE80AdFRXNN8TPB6rubxXoYXzFiydRhxvb7q/e6nsKnuvH/hiym8q48R6TBL5oh2S30St5hYqEwW+8SCMdcgigDeormrj4leFLOezin8R6ZCbyNpbd3ukEcqqSG2vnaTkNxnPyt/dOHL8RvC01veTQeIdNvEs42lnWzuUneNQQCSqEngkDp1IpN2V2VGMpyUYq7Z0dFYVt440G6MoGpwReU2xvtBMPzAZIG8DOO+Onepm8YaChcNrenKYzhwbuP5enXnjqPzqeePc3eFrp2dN/czXorIg8WaRO0C/b4YmuG224uCYjP0wY9+PMByMFcg560xfGmgPHM6a1YSRwxSTyyJcIyxxpt3sxBwANy5z60+ePcl4etFNuD08mbVFYl3400Wxu5bae+VJYo4pn+RiqpI+xG3AYwT78AEnAGae/jLQI42kbW9OCKQC32uPAJ6Dr1ODilzx7lfVq7Sag7Pyf9dTYorJj8VaQziN7+G2mKNL5F2fIl2DOW2PhgvB5xjiqrfEDwxHMIn8RaUjmPzQGvIxlOcsDnkDa2cdMc0+aL0TJlQqxTlKDSXkzoKKwG+IHhdGtw3iTSFNwhlhBv4v3qB9hZfm5G75cjvx1qOX4keFY9Pnvh4i02a1gVWkkt7pJdoZxGpIUk4LkL9eKowOjormh8TPB7NAo8V6GWuM+SP7RhzJgAnb83OAR09RXS0AFFFFABRRRQAUUUUAeF3knh6+1a6mvPh550tvqZYXcbSOxk+0vukyEGDtaV9uTn7hwOlq81DwrcyXV/N4J1KdXaUm63MS7O4MiqS/yxsZXYjIUqZsjDEN7VRXPyVP5/wPZeKwT/5hv/J5Hhmn6D4YsNcn1NvDF3CsEE9glgLp5PlYBXk8sqBl1/cgl85URqOCRq3Wn+GfFt9Zad/wjWoRD7FFZx6g0pRbaMHKKMsQXjcqRuBw6Kc5VcevUVSjPrL8DGWIwrvahr/iZxGi/BfwdoMMCWujpuigaDzZJHZnVovKctzgllzk45LE9STV7Tfhf4W0jR5NKs9Ft4NOkxut1LbTjdgdeg3sAOwOOwrqaK2PMObh+Hfh23huIk01Qk9s1m+ZHJ8k4/dglsqo2qABjaFAGMCoY/hb4Vh02OwTR4UtI7QWKRq7jbAJBIEBzkAOAw+lZHxO+PngT4PqieKPEEFjdyLujsY1aa4Ydj5aAkA+pwPeuC8N/tzfCLxFqCWba7caS8jbUk1K0eKIn3cZVR7sQK6oYTEVIc8Kba9DCVelGXLKST9T0vSfg94S0WGKK30lWWFZY4vNldzHHJv3Rrk8Jh2GB/eJ6kklr8IvDNnezTRWbi3lh8n7D5h+zrl43Zwv98tFFliScRqOgrrbS7gv7WK5tZo7m2mQSRzQsGR1IyGBHBBHcVS8R+JtJ8IaPcarreo22labbjMl1dyiNF9Bk9z2HU1zKLbslqbXSVzlLr4G+D7ix0+zi01rO2s737cqQSt+9clS6yFiSyOUTcM87R75i1H4C+DNRlglbTpIpI7hbiRluHc3GG3FJC5bcjHlh/Fjk15ze/t7fCC01A2yavf3UYODdQ6fJ5X1+YBsfhXsHw/+KPhT4qaU2o+FdbtdYtkIEghJEkRPQPGwDIT/ALQFdNTC16MeepBpeaMoVqc3aEk2Oj+GnhqLT7CyTS0S3sc/ZdsjhoTljuV924MC7ENnIJyDnFPsvhz4c023aC00uO2jaF4MQsykI7KzAEHI+ZVORyCBiukrxn4g/te/C34b6lLpuoeIhfalC22W10uFrkxkdQzL8gI7jdkelZU6NTEPkpxcvxNHVVG03K1j0M/D3w8ysp01SrRiIr5j48sHITG7hQfm29N3zdeaJvh34duLUW8umJJCAvyNI5zgkgn5uSMnk8/M39454P4cftYfDH4oalFpukeIlt9UlO2Ky1GJraSQnoELDaxP91ST7V6/WdTCujLlqQs/NHXDMK81zQrN/wDbz/zMGTwNokrxs1llojmM+dJ+79AvzfKF/hA4U8rinw+CtDt7Y2y6ZCbdrd7RoXBZGic5ZCpJBBPrWV8Rvi54Q+E2npeeK9dttIjkz5UTkvNLjrsjUFm+oGBnmvKNL/bz+EGo6gLZ9YvrFWbaLi6sJBF+agkD3IFbU8DVqLnp0m13SMamYSX7upWfo5HtWo+BtB1e7+03ulwXMnleQBICUCYIwEztHB64zwPQYrTfDXwzcSrJLpEMjqYW+YsQTEoWPIzg4AA564Gc4Fa+h67p3ibSrfU9JvrfUtOuF3w3VrKJI3HqGBwam1DUrTSLGe9vrmGys4EMktxcSBI41HVmY8Ae5rldKLdnHX0OiONxEIrkqySS7vb/ACMibwFodyxaayaZynlmSSeRnIOc5YtknBK567Tt6cVW1D4Y+GNVuLae60pJJbbb5LCWRTGVTYjLhhhlHRuoIBByAa8m179uz4Q6HqDWia3damUO1prCykeIH2ZgNw91yK9I+Gfxw8EfF+3kfwpr9vqU0S7pbQhoriMdMtG4DYzxuxj3rrlgatGPtJUml3sc39oSrfu3Vv5XuXbb4U+FLN43g0hYXiTy4mSaQNEmclUO75QSWyBgHe+c7my3SPhJ4R0GHybDQ4LaLylgEas5URrL5qqAW4AcA/gB0AFdfRXOScsPhf4WE/nf2PCZPLliDMznakgYSKMnhTvc4HGXY9WJPU0UUAFFFFABRRRQAUUUUAFFFFABRRRQAV5z+0H8VP8AhTfwl13xPGiy30Maw2Ub8hriQhUyO4BO4juFNejV85/t8eHrvXf2eL+a1RpBpt9b3syqMnywWjJ+g8wE+wJrswcI1MRThPZtGFeUoUpSjukfmVr2vaj4o1m81bVrybUNSvJDLPcztueRj3J/p2o1fw/qnh+SOPVNNvNNeQEot5A8RYDgkBgM1taSvgptLtf7Tl16PUcP55s4oWiBydhXcwJ425zjvXVa94s8DeIryBtUufF+qxxxti4up43n3l36hmKgEeWTj3HYGv1N1HFpRjp6fkfFqPNdt6nuv/BPn45ahpvi8/DjU7p7jSdQjkm01ZGz9mnRS7oueiuoc4/vKMfeNeXftgfHLUPi38UtSsI7lx4a0S4ks7C1VvkZkJV5yO7MQcHsuB65s/sgeHYdd/ai8PNoH2yTStPaa9aa7RVlWJYWGXCkgZdlXg/xCvN/EXh3TfCnxK8U6N4tGpRGyuriAfYEQyGUS/Kx3kDaVycj1FeZChRjj51Uve5U7erd366I7JVKjw0YN6Xa/I5SHQ9SuNLm1KLT7qXToW2SXaQsYUbjhnAwD8y9T/EPWuh+FXxQ1v4Q+NbDxJoVw0VxbuBNBuIjuYsjdE47qw/I4I5ANbdn4g8D2Ph+409bvxXcQSMsv2GRoo7feTEHOFfklVfkgfdQfTnPFR8HvawN4b/tpLncBLHqSxeXjHJUoSc57EdO5r0uZVU4Tjo/I5LclpRep92ftn/tKS6P8J/DFn4Tu5LW48ZWgvDdRttkhsiinAI+6zlwuR0CuOuDX566fpt3q94lrY2s17dSZKQ28ZkdsAk4UDJwAT9BX0d+1v8AD3VfCHgj4MT6hFIip4bj0+YEf6mdMSMh9DiXHvsPpXlel6l4C0bVPttjfeLrKWMMIpLdYEk+ZGVl3B8qDkDIzwW4rzctp06GGXsle7evzsdeLlKpW9/S1vyOEvLO60u8e3uoJrO6iI3RTIUdD1GQeR2r9Ff2Wf2oZdW+APii/wDFM73+qeCrffLM7fvLuAoxhyT1csjR574UnkmvinWNV+HetXmpXNx/wk32uWaRorhfJfzF3NtZwzZ3Ebc88c9a9Q/Zt+HureJv2f8A443tpFIYpNPt4YQoJ86SFmuJFX1IULx/tilmNOniMOvbK1nH8Wk/wHhZTpVfcd7p/keGfEL4ga38UPFt/wCItfu2u9Qu3LHJOyJc/LGg/hRRwB/Wsu+0DVNNs7e7vNNu7S0uADDPPAyJKCAwKsRg5BB47GtTw+PCJ08f24+sre+f/wAw+OIx+Tgf32B3Z3cYx05rrr7xV4I1DTdM0+9vPF2pWVv8rrPLGDGFSNU8tN5UADzFHoNvqRXqc3s7QhHReRxW5ryk9Tvf2JPjlqHw3+KWneG7m6d/DXiG4W0ltnbKQ3D4WKVR2JbareoPP3Rjo/2+vjlqHij4gz+ArC5eHw/opQXUcbYF1dFQxLeoQEKAejBj6Y8c+F/h3T/FPx28Fab4TGoyW82p2sh+3ogljCSB5G+QkbVRSc9eDW9+0p4etfDn7TnjCHxKL1NOub6S932aq0rRyp5kZTcQCAWAPP8ACR1rzHQovHqrb3uW9vna/qdiqVPqzhfS9v8AgHkGn6HqOrQzy2Wn3V5Fb7fOkt4WdY85xuIHGcHGfQ1Z8K+KtV8EeIbHXNEvZdP1SykEsNxCcEEdj6gjgg8EEg8V23h3xP4O8NWt6LHVPGFtJcRNugheKGKR1WXy97I+SMsoxjjc/XvgeIn8DS6SW0RNdg1TIxHeCFoDlueQ24YXpwcn0r0+fmbjKOj8vzOPlslJPU/W74L/ABHi+LXwv8PeK4o1hfULfM8K9I5lYpKo9g6tj2xXa14V+xL4eu/Dv7N/hdLxGjlvPPvUjYciOSVih+hXa3/Aq91r8qxUI0684Q2TdvvPtaMnKnGUt2kFFFFcpsFFFFABRRRQAUUUUAFFFFABRRRQAVV1TTLTWtNutPv7eO7srqJoJ4JRlJI2BDKR3BBIq1WV4p8Tad4M8OalrurXAtdN0+BrieU84VRk4HcnoB3JAqo3bSjuJ2tqfn18bP8Agn/4r8O6xc33gFF8RaFIxeOxaZUu7Yf3TvIEgHYg7j3Hc+Z+G/2N/i74k1BLVfCNxpqFsPc6lIkEUY9Tk5I/3QT7VpfGz9snx18VNYuY9M1S78L+HAxW30/T5jFIydjNIpDMx7jO0dh3Pmfhv4xeOfCGoJe6R4t1iyuFbd8t47I3syMSrD2YEV+mYeOYqilUlHm80/xsz5CpLC+091Ox+nH7M/7NOl/s+eHZx566p4k1AL9u1EJtXA5EUYPIQH15Y8nHAHE/tWfsdxfGq6Pifw1cQab4tWMJPHcZWC+VRhdxAO1wBgNgggAHGARo/siftTH47aZc6NrqQ23i/TYhLIYRtjvIchfNVf4WBIDL0+YEcHA8/wD2wv2x9R8C63ceB/A06W+qwKBqWrbQ5gYjPlRA5G7BG5jnGcDkEj5OlTzD+0Gk/wB51fS3+R7c5YX6qr/D+N/8z5bvf2Rfi/Y6gbN/A9/JJnAkheOSI++9WK4/Gvo79mv9g2/0TxBZeJ/iOLdDZus1toMMgm3SA5Vp3GVwDzsUkHjJxkH43vfid4w1LUDfXXivWri9J3faJNQlL59ju4r6O/Zq/be8ReFPEFjoXj3Updc8N3LrD/aN42+5sSTgOZDzImfvBskDkHjB+nx0MwlQapyV+tk0/lds8fDywqqLmT+e3zPuj4vfCfRPjR4HvPDWuowglIkguY8eZbTAHbIme4yRjuCR3r85viF+w38UvBepSx6fpA8U6buPlXulupLDtuiYh1PrwR7mvv8A/aC+OmmfAf4fya9cxrfX1w3kadYhsfaJiCRk9kAGSfoOpFfmN8Qv2iviH8TNSlutY8UX6ROxKWNlM1vbRDsFjUgcepyfUmvGySON5W6bSh59/I78wlh7pTXveX6noXw3/YU+JnjTVIV1jTl8JaVuHm3moOpk299kKksW/wB7aPev0c+GXw10T4T+CrDwxoUBjsLVTl5MGSeQ8vI57sx/DoBgACvyp+G/7S/xF+F+qQ3Ol+Jb27tVYGTTdRma4tpR3BRj8ufVSD71+nfwV+NmjfGb4bQeLLQrYLGGj1C3mcH7HKgBdWb+7ghg3GVIPHIE53HG2TqtOHl38ysulh9VBe95nyf+0R+wJqsmuXmv/DVILmyuXMsmgySLE8DHk+SzEKU6/KSCOgz28F0v9kL4v6rqC2ieCL6Bs4Mt08cUS++9mAI+ma9B/aI/bc8U+PNcvNL8F6lceHPC0LmOOezYxXV4Bx5jSD5kU9lXHB5z28E0v4peMtF1AX1j4r1q1uw27zor+UMT7/Nz+Ne7g45gqCVSUb9Lpt/OzR5teWFdR8idvL9D9G/2Vf2R7X4ErJrut3MOq+L7mLyvMhBMFnGfvJGSAWY93IHHAAGS2v8AtQ/st6d+0DpVveWtxHpXiuwjMdrfSKTHLHknyZcc7ckkMMlSTwckV53+x3+1/e/FDUV8F+M5I38ReWz2OpIoT7aFGWR1HAkABYEAAgHgEfNuftg/tYS/BWOHwz4YEMvi28h86S4lUOlhEchW2nhpGwcA8ADJByAflpU8w/tC1/3nfpb/AC/rc9lSwv1X+5+N/wDM+NNe/Y8+L2g6g1q/g26vgDhbiwkjmiceoKtwP94A+1evfAn/AIJ/+IdW1q11T4jImjaNC4kOkxzLJc3WOQrFCVjQ9+d3UYHUfM+vfFrxt4m1Br7VPFms3tyx3b5L6TC/7oBwo9gAK9e+BP7aXjT4Y61a23iDUrvxT4Xdwk9vfSGW4hX+/FIx3ZH90nacY4zkfV4mOYui1TlHm8k7/K7PEoywqqe8nb+tz9QbW1hsraK3t4kgghQRxxRqFVFAwFAHQAVLVLRdYsvEWj2WqadcJd6fewpcW88f3ZI3UMrD6girtfmbvfU+v9AooopDCiiigAooooAKKKKACiiigAooooAK+bP+CgeoXdl+zxcxWxYRXep2sNxt/wCeeWfn23olfSdcX8Y/hrafF74ba54UvH8gX8OIbjGfJmUho3+gZRkdxkd668JUjRxEKk9k0YV4OpSlGO7R+RGmfDXxNrGmW2o2Wj3FzZXAdopU24cKSGxzngqfyruPGnwi1HUL20j8O+DZtEIjd5ba61WGWVhvcKSGcEEbGU4AGRjGRk8T4+8B+JPhX4mudA8Q2c+m39uxABJ2SryA8bdGQjOCPp6iufa+uWZSbiUlRtBLngZJx+ZP51+p+/UtOElbpv8A56nxnuxvGS1/ryPcP2R4dX8I/tS+FLCaCW0vfOmt7m2cYOxrdywYeww34CvN9X0fX/iF4+8Sy29nNqGqNdXN7dKpG4ZlO8nJ/vMPzr7C/YL/AGdNW0fVH+JPie0lsnaBodItrlSJWDjD3BB5AK5Vc9QzHptJ8u/bS/Z11f4c+O9T8ZaPaSz+FNYna5kmt1J+xTucuj4+6rMSVPT5tvUc+TDGUpY+VNNc3KlfpdN3X4nbLDzWGU2tLt/LQ830P4V3i+DtSTUPBd3JqzEm21STUo7eGIMYVQbGIDYZ+Tn+MDg4I4nxN8PfEPg63huNY017S3mYLHLvR0LEZ25UkZx26isMX1yEKC4lCMApXecEDGB+g/IV6X8Cfgf4j+PnjC20uyS4TR4ZFbUNUcExWsffk8FyOFXqT7AkepKToJ1Kkly9d/8AM40lUahBa/15HbftUeINV1rwT8Eor9pHjXwnDMpbPzSMQjN7krHGfxrz3w98JfENjr6f2z4RvdRs4lfzrWK4WEkmNypMmTtUEBiemFIr72/a6/Zpb4mfC7Ro/Clqo1nwrF5dhZqcedahFVoFP94BEK5/ukfxZr8z7xb/AEu8ntLoXNpdQuY5YJtyOjDIKsp5BHPB9687La8cTh+Wk0mr3Xq790deLpujVvNX2/I7fXPg34qa/wBRnsvDvkWcMzp9nt7uOZotrFSpHmFiRtOfT2GK9F/Z58QarpfwE+PFvZPJ5DaXaFtucLvkeJ8emY2fPsvtXhGmw6rr2pRWNgl3f393Jtjt4N0kkzk9ABySc1+mv7NP7MMPw/8AgjrHh/xRCr6r4qiYatHGwJhjZCiQhhwSgZjkdGY4yADSzLERw1BKs022rL0ab6vsPCUnWqXgraP8Ufm1oXgDxD4m0/7dpmlzXlp5/wBm82MrjzMA7eT6MPzrvNb+Et9feHdFg0vwZc6XrDgCee81SLNxtjj3lY2YbQWkDYxwrDJweMX4yfBvxN8C/F0+jaxFMtvvLWWoxgiG7jB+V0Pr0yvVT+BPB/brkqgNxKQhJUbzxnGcfkPyFemnKslUpyVum/8AmcelNuM1r/Xkd/8ADnS9e+H/AMbvBSXFpNYaomr2ckcbYy6tMq8YOCrcj3Brf/aMsta8cftNeNrOC3kvdSfUpYYYQRuMcSYXGT0EaCvVP2Hf2ddX8W+NtP8AiDr1rNb+HtJfz7BrgEG9uB9xkz1RD827puAAz82Oh/by/Zz1ZfEk3xK8OWkt5ZXMajV4bdSXt5EUKJ8D+AqFDEdCuT97jynjKSx6pNrm5bX6Xvt/T8jsWHn9W57aXvby7nz54K+F97b2eqPr3gm+1IeV5kFwL9LWO3CJKz7snDE7AevARjg1yWv/AAv8T+GNNOo6jpbQ2KkA3Ec0cqDLbRkoxxk9M9e1c4t9coCFuJVBBUgORwc5H45P5mur+Gfw18T/ABi8UW3h7w/bTXk8jDzZGJ8m2jzzJI3RVH5noMkgV6z5qbdSckl13/z0OJWnaMVr/XkfpH+wzqF3qH7Nfhr7UWcQS3UMLN1Mazvj8Bkj8K98rmfhr4Dsfhj4D0PwtpxL2umWywiQjBkbq8hHYsxZj9a6avyrFVI1a86kdm2/xPtKMXCnGL3SCiiiuY2CiiigAooooAKKKKACiiigAooooAKKKKAMPxZ4H8PeO7AWXiLRLDW7VTlY763WUIfVcj5T7jmuX8N/s7/DTwjqCX+leCdHtr2Nt0c7W4keM+ql87T7jFeiUVrGtUjHljJpepDpwk7tahUc0Md1DJDNGssUilXjdQysDwQQeoqSisizzC9/Zj+FOoagb2bwFovnk7j5dsI0J/3Fwv6V6Douhab4b02LT9J0+10uwiGI7WzhWKJPoqgAVeorWdapUVpybXmyIwjF3irBXG+NPg34H+Ik4uPEnhXS9WugMfaZ7dfOx6eYMNj2zXZUVMZypvmg7McoqStJXOR8F/CPwX8OWZ/DXhjTNGmYbWuLa3USsPQyH5iPbNddRRRKcpvmk7sIxUVaKsZviDw3pPizTZNO1rTLPV7CTlra+gWaM++1gRmuE0v9mf4V6PqC3tr4D0UXCncplthKqn1CtlR+Ar02iqjWqU1ywk0vUUqcJO8lcbHGsaKiKERRgKowAPQUtLRWRZ5vr37N/wAMPEuoNfah4H0aW7c7nkjthEXPq2zG4+5rsPC/g/QvBOmiw8P6PY6LZZz5FjbrCpPqQoGT7nmtiitZVqk48spNr1IVOEXdLUKKKKyLCiiigBskiRLudlQerHFRfbrf/nvF/wB9iodWs5b62EcMxt5N2fMXOV4I4/OuTurfXNNkWKK81aZ41ZnlhgililBZsD5yWDAOOB/zz54oHbS9zsvt1v8A894v++xT4545s+XIr467SDXC2s3iK4mW2+161AWUoLiawtfLXJUhzyTkDcMYx6jI57axt5be3WOad7pwWzLIFDHJJHCgDgcfhQI//9k=";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion
    }
}
