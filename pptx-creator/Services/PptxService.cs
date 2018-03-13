using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace pptx_creator.Services
{
    public class PptxService
    {
        /*static void Main(string[] args)
        {
            string filepath = @"C:\Users\username\Documents\PresentationFromFilename.pptx";
            CreatePresentation(filepath);
        }*/

        //public static void CreatePresentation(string filepath)
        public static void CreatePresentation()
        {
            string filepath = @"my_pptx.pptx";

            // Create a presentation at a specified file path. The presentation document type is pptx, by default.
            PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            CreatePresentationParts(presentationPart);            

            //Close the presentation handle
            presentationDoc.Close();
        }

        public static void InsertSlide()
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(@"my_pptx.pptx", true))
            {
                int position = 1;
                string slideTitle = "My new slide.";
                // Pass the source document and the position and title of the slide to be inserted to the next methoDrawing.
                InsertNewSlide(presentationDocument, position, slideTitle);
            }
        }

        // Insert the specified slide into the presentation at the specified position.
        public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        {

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slideTitle == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));

            // Declare and instantiate the body shape of the new slide.
            Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
            drawingObjectId++;

            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the body shape.
            bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
                    new Drawing.ListStyle(),
                    new Drawing.Paragraph());

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        private static void CreatePresentationParts(PresentationPart presentationPart)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

           presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

           SlidePart slidePart1;
           SlideLayoutPart slideLayoutPart1;
           SlideMasterPart slideMasterPart1;
           ThemePart themePart1;

            
            slidePart1 = CreateSlidePart(presentationPart);
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1); 
  
            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");            
        }

    private static SlidePart CreateSlidePart(PresentationPart presentationPart)        
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
                slidePart1.Slide = new Slide(
                        new CommonSlideData(
                            new ShapeTree(
                                new P.NonVisualGroupShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                    new P.NonVisualGroupShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()),
                                new GroupShapeProperties(new Drawing.TransformGroup()),
                                new P.Shape(
                                    new P.NonVisualShapeProperties(
                                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                        new P.NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                    new P.ShapeProperties(),
                                    new P.TextBody(
                                        new Drawing.BodyProperties(),
                                        new Drawing.ListStyle(),
                                        new Drawing.Paragraph(new Drawing.EndParagraphRunProperties() { Language = "en-US" }))))),
                        new ColorMapOverride(new Drawing.MasterColorMapping()));
                return slidePart1;
         } 
   
      private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new Drawing.TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                            new P.NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.EndParagraphRunProperties()))))),
                new ColorMapOverride(new Drawing.MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
         }

   private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
   {
       SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
       SlideMaster slideMaster = new SlideMaster(
       new CommonSlideData(new ShapeTree(
         new P.NonVisualGroupShapeProperties(
         new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
         new P.NonVisualGroupShapeDrawingProperties(),
         new ApplicationNonVisualDrawingProperties()),
         new GroupShapeProperties(new Drawing.TransformGroup()),
         new P.Shape(
         new P.NonVisualShapeProperties(
           new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
           new P.NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
           new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
         new P.ShapeProperties(),
         new P.TextBody(
           new Drawing.BodyProperties(),
           new Drawing.ListStyle(),
           new Drawing.Paragraph())))),
       new P.ColorMap() { Background1 = Drawing.ColorSchemeIndexValues.Light1, Text1 = Drawing.ColorSchemeIndexValues.Dark1, Background2 = Drawing.ColorSchemeIndexValues.Light2, Text2 = Drawing.ColorSchemeIndexValues.Dark2, Accent1 = Drawing.ColorSchemeIndexValues.Accent1, Accent2 = Drawing.ColorSchemeIndexValues.Accent2, Accent3 = Drawing.ColorSchemeIndexValues.Accent3, Accent4 = Drawing.ColorSchemeIndexValues.Accent4, Accent5 = Drawing.ColorSchemeIndexValues.Accent5, Accent6 = Drawing.ColorSchemeIndexValues.Accent6, Hyperlink = Drawing.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = Drawing.ColorSchemeIndexValues.FollowedHyperlink },
       new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
       new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
       slideMasterPart1.SlideMaster = slideMaster;

       return slideMasterPart1;
    }

   private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
   {
       ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
       Drawing.Theme theme1 = new Drawing.Theme() { Name = "Office Theme" };

       Drawing.ThemeElements themeElements1 = new Drawing.ThemeElements(
       new Drawing.ColorScheme(
         new Drawing.Dark1Color(new Drawing.SystemColor() { Val = Drawing.SystemColorValues.WindowText, LastColor = "000000" }),
         new Drawing.Light1Color(new Drawing.SystemColor() { Val = Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
         new Drawing.Dark2Color(new Drawing.RgbColorModelHex() { Val = "1F497D" }),
         new Drawing.Light2Color(new Drawing.RgbColorModelHex() { Val = "EEECE1" }),
         new Drawing.Accent1Color(new Drawing.RgbColorModelHex() { Val = "4F81BD" }),
         new Drawing.Accent2Color(new Drawing.RgbColorModelHex() { Val = "C0504D" }),
         new Drawing.Accent3Color(new Drawing.RgbColorModelHex() { Val = "9BBB59" }),
         new Drawing.Accent4Color(new Drawing.RgbColorModelHex() { Val = "8064A2" }),
         new Drawing.Accent5Color(new Drawing.RgbColorModelHex() { Val = "4BACC6" }),
         new Drawing.Accent6Color(new Drawing.RgbColorModelHex() { Val = "F79646" }),
         new Drawing.Hyperlink(new Drawing.RgbColorModelHex() { Val = "0000FF" }),
         new Drawing.FollowedHyperlinkColor(new Drawing.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
         new Drawing.FontScheme(
         new Drawing.MajorFont(
         new Drawing.LatinFont() { Typeface = "Calibri" },
         new Drawing.EastAsianFont() { Typeface = "" },
         new Drawing.ComplexScriptFont() { Typeface = "" }),
         new Drawing.MinorFont(
         new Drawing.LatinFont() { Typeface = "Calibri" },
         new Drawing.EastAsianFont() { Typeface = "" },
         new Drawing.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
         new Drawing.FormatScheme(
         new Drawing.FillStyleList(
         new Drawing.SolidFill(new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.PhColor }),
         new Drawing.GradientFill(
           new Drawing.GradientStopList(
           new Drawing.GradientStop(new Drawing.SchemeColor(new Drawing.Tint() { Val = 50000 },
             new Drawing.SaturationModulation() { Val = 300000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 0 },
           new Drawing.GradientStop(new Drawing.SchemeColor(new Drawing.Tint() { Val = 37000 },
            new Drawing.SaturationModulation() { Val = 300000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 35000 },
           new Drawing.GradientStop(new Drawing.SchemeColor(new Drawing.Tint() { Val = 15000 },
            new Drawing.SaturationModulation() { Val = 350000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 100000 }
           ),
           new Drawing.LinearGradientFill() { Angle = 16200000, Scaled = true }),
         new Drawing.NoFill(),
         new Drawing.PatternFill(),
         new Drawing.GroupFill()),
         new Drawing.LineStyleList(
         new Drawing.Outline(
           new Drawing.SolidFill(
           new Drawing.SchemeColor(
             new Drawing.Shade() { Val = 95000 },
             new Drawing.SaturationModulation() { Val = 105000 }) { Val = Drawing.SchemeColorValues.PhColor }),
           new Drawing.PresetDash() { Val = Drawing.PresetLineDashValues.Solid })
         {
             Width = 9525,
             CapType = Drawing.LineCapValues.Flat,
             CompoundLineType = Drawing.CompoundLineValues.Single,
             Alignment = Drawing.PenAlignmentValues.Center
         },
         new Drawing.Outline(
           new Drawing.SolidFill(
           new Drawing.SchemeColor(
             new Drawing.Shade() { Val = 95000 },
             new Drawing.SaturationModulation() { Val = 105000 }) { Val = Drawing.SchemeColorValues.PhColor }),
           new Drawing.PresetDash() { Val = Drawing.PresetLineDashValues.Solid })
         {
             Width = 9525,
             CapType = Drawing.LineCapValues.Flat,
             CompoundLineType = Drawing.CompoundLineValues.Single,
             Alignment = Drawing.PenAlignmentValues.Center
         },
         new Drawing.Outline(
           new Drawing.SolidFill(
           new Drawing.SchemeColor(
             new Drawing.Shade() { Val = 95000 },
             new Drawing.SaturationModulation() { Val = 105000 }) { Val = Drawing.SchemeColorValues.PhColor }),
           new Drawing.PresetDash() { Val = Drawing.PresetLineDashValues.Solid })
         {
             Width = 9525,
             CapType = Drawing.LineCapValues.Flat,
             CompoundLineType = Drawing.CompoundLineValues.Single,
             Alignment = Drawing.PenAlignmentValues.Center
         }),
         new Drawing.EffectStyleList(
         new Drawing.EffectStyle(
           new Drawing.EffectList(
           new Drawing.OuterShadow(
             new Drawing.RgbColorModelHex(
             new Drawing.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
         new Drawing.EffectStyle(
           new Drawing.EffectList(
           new Drawing.OuterShadow(
             new Drawing.RgbColorModelHex(
             new Drawing.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
         new Drawing.EffectStyle(
           new Drawing.EffectList(
           new Drawing.OuterShadow(
             new Drawing.RgbColorModelHex(
             new Drawing.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
         new Drawing.BackgroundFillStyleList(
         new Drawing.SolidFill(new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.PhColor }),
         new Drawing.GradientFill(
           new Drawing.GradientStopList(
           new Drawing.GradientStop(
             new Drawing.SchemeColor(new Drawing.Tint() { Val = 50000 },
               new Drawing.SaturationModulation() { Val = 300000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 0 },
           new Drawing.GradientStop(
             new Drawing.SchemeColor(new Drawing.Tint() { Val = 50000 },
               new Drawing.SaturationModulation() { Val = 300000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 0 },
           new Drawing.GradientStop(
             new Drawing.SchemeColor(new Drawing.Tint() { Val = 50000 },
               new Drawing.SaturationModulation() { Val = 300000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 0 }),
           new Drawing.LinearGradientFill() { Angle = 16200000, Scaled = true }),
         new Drawing.GradientFill(
           new Drawing.GradientStopList(
           new Drawing.GradientStop(
             new Drawing.SchemeColor(new Drawing.Tint() { Val = 50000 },
               new Drawing.SaturationModulation() { Val = 300000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 0 },
           new Drawing.GradientStop(
             new Drawing.SchemeColor(new Drawing.Tint() { Val = 50000 },
               new Drawing.SaturationModulation() { Val = 300000 }) { Val = Drawing.SchemeColorValues.PhColor }) { Position = 0 }),
           new Drawing.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

       theme1.Append(themeElements1);
       theme1.Append(new Drawing.ObjectDefaults());
       theme1.Append(new Drawing.ExtraColorSchemeList());

       themePart1.Theme = theme1;
       return themePart1;

         }
    } 
} 
