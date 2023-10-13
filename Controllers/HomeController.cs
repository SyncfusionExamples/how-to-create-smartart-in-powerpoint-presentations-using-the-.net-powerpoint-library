using Microsoft.AspNetCore.Mvc;
using PPTXSmartArt.Models;
using System.Diagnostics;
using Syncfusion.Presentation;

namespace PPTXSmartArt.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult CreateSmartArt()
        {
            IPresentation presentation = Presentation.Create();
            CreateSlideOne(presentation);
            CreateSlideTwo(presentation);
            CreateSlideThree(presentation);
            CreateSlideFour(presentation);

            MemoryStream stream = new MemoryStream();
            presentation.Save(stream);
            return File(stream, "application/presentation", "SmartArt.pptx");
        }

        private void CreateSlideFour(IPresentation presentation)
        {
            ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
            IShape textbox = slide.AddTextBox(250, 50, 550, 50);
            IParagraph paragraph = textbox.TextBody.AddParagraph("3 Step Writing Process");
            paragraph.Font.FontName = "Calibri";
            paragraph.Font.FontSize = 40;
            paragraph.Font.Color = ColorObject.OrangeRed;

            ISmartArt smartArt = slide.Shapes.AddSmartArt(SmartArtType.StepUpProcess, 100, 100, 620, 407);
            ISmartArtNode nodeOne = smartArt.Nodes[0];
            nodeOne.TextBody.AddParagraph("Prewriting");
            ISmartArtNode nodeTwo = smartArt.Nodes[1];
            nodeTwo.TextBody.AddParagraph("Writing");
            ISmartArtNode nodeThree = smartArt.Nodes[2];
            nodeThree.TextBody.AddParagraph("Revising");
        }

        private void CreateSlideThree(IPresentation presentation)
        {
            ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
            IShape textbox = slide.AddTextBox(300, 50, 550, 50);
            IParagraph paragraph = textbox.TextBody.AddParagraph("Organization Chart");
            paragraph.Font.FontName = "Calibri";
            paragraph.Font.FontSize = 40;
            paragraph.Font.Color = ColorObject.OrangeRed;

            ISmartArt smartArt = slide.Shapes.AddSmartArt(SmartArtType.Hierarchy, 100, 120, 620, 407);
            ISmartArtNode nodeOne = smartArt.Nodes[0];
            nodeOne.TextBody.AddParagraph("CEO");
            ISmartArtNode childOne = nodeOne.ChildNodes[0];
            childOne.TextBody.AddParagraph("Vice President Finance");
            ISmartArtNode childTwo = nodeOne.ChildNodes[1];
            childTwo.TextBody.AddParagraph("Vice President Production");
            ISmartArtNode childOneOfOne = childOne.ChildNodes[0];
            childOneOfOne.TextBody.AddParagraph("Accountant");
            ISmartArtNode childTwoOfOne = childOne.ChildNodes[1];
            childTwoOfOne.TextBody.AddParagraph("Budget Analyst");
            ISmartArtNode childOneOfTwo = childTwo.ChildNodes[0];
            childOneOfTwo.TextBody.AddParagraph("Production Manager");
        }

        private void CreateSlideTwo(IPresentation presentation)
        {
            ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
            IShape textbox = slide.AddTextBox(150, 50, 550, 50);
            IParagraph paragraph = textbox.TextBody.AddParagraph("Software Development Life Cycle");
            paragraph.Font.FontName = "Calibri";
            paragraph.Font.FontSize = 40;
            paragraph.Font.Color = ColorObject.OrangeRed;

            ISmartArt smartart = slide.Shapes.AddSmartArt(SmartArtType.BasicCycle, 100, 120, 620, 407);
            ISmartArtNode nodeOne = smartart.Nodes[0];
            nodeOne.TextBody.AddParagraph("Analysis");
            ISmartArtNode nodeTwo = smartart.Nodes[1];
            nodeTwo.TextBody.AddParagraph("Design");
            ISmartArtNode nodeThree = smartart.Nodes[2];
            nodeThree.TextBody.AddParagraph("Implementation");
            ISmartArtNode nodeFour = smartart.Nodes[3];
            nodeFour.TextBody.AddParagraph("Testing");
            ISmartArtNode nodeFive = smartart.Nodes[4];
            nodeFive.TextBody.AddParagraph("Evaluation");
        }

        private void CreateSlideOne(IPresentation presentation)
        {
            ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
            IShape textbox = slide.AddTextBox(270, 50, 400, 50);
            IParagraph paragraph = textbox.TextBody.AddParagraph("5 P's of Marketing");
            paragraph.Font.FontName = "Calibri";
            paragraph.Font.FontSize = 40;
            paragraph.Font.Color = ColorObject.OrangeRed;

            ISmartArt smartArt = slide.Shapes.AddSmartArt(SmartArtType.BasicBlockList, 100, 120, 620, 407);
            ISmartArtNode nodeOne = smartArt.Nodes[0];
            nodeOne.TextBody.AddParagraph("Product");
            ISmartArtNode nodetwo = smartArt.Nodes[1];
            nodetwo.TextBody.AddParagraph("Price");
            ISmartArtNode nodeThree = smartArt.Nodes[2];
            nodeThree.TextBody.AddParagraph("Promotion");
            ISmartArtNode nodeFour = smartArt.Nodes[3];
            nodeFour.TextBody.AddParagraph("Place");
            ISmartArtNode nodeFive = smartArt.Nodes[4];
            nodeFive.TextBody.AddParagraph("People");
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}