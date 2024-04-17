using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Input;
using Microsoft.Windows.SemanticSearch;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Controls;

using static System.Net.Mime.MediaTypeNames;
using Microsoft.Windows.Imaging;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Graphics.Imaging;

namespace XpsViewer
{
    internal class MyDocumentViewer : DocumentViewer
    {
        //Semantic Search Data
        public List<BitmapSource> bitmapSource;
        public List<System.Windows.Shapes.Path> paths;
        public List<System.Windows.Controls.Canvas> canvases;

        public string docText;
        public string[] docSentences;
        public List<string> actualSentences;
        public EmbeddingVector[] embeddings;
        public EmbeddingVector[] imageEmbeddings;

        public int[] indexes;
        public float[] scores;
        public int[] rank;
        public int currentIndex = 0;
        public bool updated = false;
        ImageSearchEmbeddingsCreator embedCreator;
        ImageSearchEmbeddingsCreator imageEmbedCreator;
        String currentSearchString;
        IList textPointersStart;
        IList textPointersEnd;
        object searchTextSegment;

        bool imageHighligted = false;
        System.Windows.Shapes.Rectangle highlightRect;
        private ToolBar _myfindToolbar; // MS.Internal.Documents.FindToolBar
        private object _mydocumentScrollInfo; // MS.Internal.Documents.DocumentGrid

        private MethodInfo _miFind; // DocumentViewerBase.Find(FindToolBar)
        private MethodInfo _miGoToTextBox; // FindToolBar.GoToTextBox()
        private MethodInfo _miMakeSelectionVisible; // DocumentGrid.MakeSelectionVisible()

        /// <summary>
        /// Limit for returned search results. 0 for no limit, default is int.MaxValue.
        /// </summary>
        public int MaxSearchResults { get { return (int)GetValue(MaxSearchResultsProperty); } set { SetValue(MaxSearchResultsProperty, value); } }
        public static readonly DependencyProperty MaxSearchResultsProperty =
            DependencyProperty.Register("MaxSearchResults", typeof(int), typeof(MyDocumentViewer), new PropertyMetadata(int.MaxValue));


        /// <summary>
        /// Determines if the search of the find toolbox is overridden and multiple search results are selected in the document.
        /// </summary>
        public bool IsMultiSearchEnabled { get { return (bool)GetValue(IsMultiSearchEnabledProperty); } set { SetValue(IsMultiSearchEnabledProperty, value); } }
        public static readonly DependencyProperty IsMultiSearchEnabledProperty =
            DependencyProperty.Register("IsMultiSearchEnabled", typeof(bool), typeof(MyDocumentViewer), new PropertyMetadata(true));


        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            if (IsMultiSearchEnabled)
            {
                // get some private fields from the base class DocumentViewer
                _myfindToolbar = this.GetType().GetPrivateFieldOfBase("_findToolbar").GetValue(this) as ToolBar;
                _mydocumentScrollInfo = this.GetType().GetPrivateFieldOfBase("_documentScrollInfo").GetValue(this);
                // Increase the size of the toolbar
                _myfindToolbar.Height = 40;
                _myfindToolbar.Width = 3000;
                Thickness margin = _myfindToolbar.Margin;
                margin.Left = 50;
                _myfindToolbar.Margin = margin;
                // replace button click handler of find toolbar
                EventInfo evt = _myfindToolbar.GetType().GetEvent("FindClicked");
                ReflectionHelper.RemoveEventHandler(_myfindToolbar, evt.Name); // remove existing handler
                evt.AddEventHandler(_myfindToolbar, new EventHandler(OnFindInvoked)); // attach own handler
                System.Windows.Controls.Button findNextButton = _myfindToolbar.FindName("FindNextButton") as System.Windows.Controls.Button;
                System.Windows.Controls.Button findPreviousButton = _myfindToolbar.FindName("FindPreviousButton") as System.Windows.Controls.Button;

                // Attach Click event handlers to the buttons
                findNextButton.Click += (s, args) => { /* Handle Find Next button click */ };
                findPreviousButton.Click += (s, args) => { /* Handle Find Previous button click */ };

                // get some methods that will need to be invoked
                _miFind = this.GetType().GetMethod("Find", BindingFlags.NonPublic | BindingFlags.Instance);
                _miGoToTextBox = _myfindToolbar.GetType().GetMethod("GoToTextBox");
                _miMakeSelectionVisible = _mydocumentScrollInfo.GetType().GetMethod("MakeSelectionVisible");
            }
        }


        /// <summary>
        /// This is replacing DocumentViewer.OnFindInvoked(object sender, EventArgs e)
        /// </summary>
        private void OnFindInvoked(object sender, EventArgs e)
        {
            var toolbarType = _myfindToolbar.GetType();
            var nextButtonField = toolbarType.GetField("nextButton", BindingFlags.NonPublic | BindingFlags.Instance);
            var previousButtonField = toolbarType.GetField("previousButton", BindingFlags.NonPublic | BindingFlags.Instance);

            if (nextButtonField != null && previousButtonField != null)
            {
                var nextButton = nextButtonField.GetValue(_myfindToolbar) as System.Windows.Controls.Button;
                var previousButton = previousButtonField.GetValue(_myfindToolbar) as System.Windows.Controls.Button;

                // Assuming the buttons have a "IsClicked" property
                var isNextButtonClicked = (bool)nextButton.GetType().GetProperty("IsClicked").GetValue(nextButton);
                var isPreviousButtonClicked = (bool)previousButton.GetType().GetProperty("IsClicked").GetValue(previousButton);

                if (isNextButtonClicked)
                {
                    // Next button was clicked
                }
                else if (isPreviousButtonClicked)
                {
                    // Previous button was clicked
                }
            }
            IList allSegments = null; // collection of text segments
            System.Windows.Documents.TextRange findResult = null; // could also use object, does not need type

            // Drill down to the list of selected text segments: DocumentViewer.TextEditor.Selection.TextSegments
            object textEditor = this.GetType().GetProperty("TextEditor", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(this); // System.Windows.Documents.TextEditor
            object selection = textEditor.GetType().GetProperty("Selection", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(textEditor); // System.Windows.Documents.TextSelection
            FieldInfo fiTextSegments = selection.GetType().GetPrivateFieldOfBase("_textSegments");
            IList textSegments = fiTextSegments.GetValue(selection) as IList; // List<System.Windows.Documents.TextSegment>

            // Get the find toolbar
            ToolBar findToolbar = this.GetType().GetPrivateFieldOfBase("_findToolbar").GetValue(this) as ToolBar;

            // Get the SearchText property
            PropertyInfo piSearchText = findToolbar.GetType().GetProperty("SearchText", BindingFlags.Public | BindingFlags.Instance);

            // Get the searchable text
            string findSearchText = (string)piSearchText.GetValue(findToolbar);

            // Get the TextContainer from the TextEditor
            FieldInfo piTextContainer = textEditor.GetType().GetPrivateFieldOfBase("_textContainer");
            object textContainer = piTextContainer.GetValue(textEditor);
            // Get the Start and End properties of the ITextContainer

            // Get the Start and End properties of the ITextContainer
            FieldInfo startProp = textContainer.GetType().GetPrivateFieldOfBase("_start");
            FieldInfo endProp = textContainer.GetType().GetPrivateFieldOfBase("_end");

            // Get the Start and End ITextPointers
            object start = startProp.GetValue(textContainer);
            object end = endProp.GetValue(textContainer);

            // Get the Start TextPointer
            object pointer = startProp.GetValue(textContainer);

            if (imageHighligted)
            {
                if (highlightRect != null)
                {
                    canvases[rank[currentIndex] - embeddings.Length].Children.Remove(highlightRect);
                }
            }

            if (currentSearchString == findSearchText)
            {
                currentIndex++;
            }
            else
            {
                currentSearchString = findSearchText;
                currentIndex = 0;

                var inputStringEmbed = embedCreator.CreateVectorForText(currentSearchString);
                //var op = ImageSearchEmbeddingsCreator.MakeAvailableAsync();

                var model8 = ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Text).AsTask().Result;
                var op = ImageSearchEmbeddingsCreator.MakeAvailableAsync();
                var embedImageText = model8.CreateVectorForText(currentSearchString);
                //rank = embedCreator.CalculateRanking();

                float[] scores = new float[embeddings.Length];
                rank = new int[embeddings.Length];

                for (int i = 0; i < embeddings.Length; i++)
                {
                    var score = CosineSimilarity(embeddings[i], inputStringEmbed);
                    scores[i] = score;
                }

                float[] scoresImage = new float[imageEmbeddings.Length];
                //rankImage = new int[embeddings.Length];

                for (int i = 0; i < imageEmbeddings.Length; i++)
                {
                    var score = CosineSimilarity(imageEmbeddings[i], embedImageText);
                    scoresImage[i] = score;
                }

                // Append scores list from image embeddings to text embeddings
                scores = scores.Concat(scoresImage).ToArray();

                var indexedFloats = scores.Select((value, index) => new { Value = value, Index = index })
                  .ToArray();

                // Sort the indexed floats by value in descending order
                Array.Sort(indexedFloats, (a, b) => b.Value.CompareTo(a.Value));

                // Extract the top k indices
                rank = indexedFloats.Select(item => item.Index).ToArray();

                int offset = 0;

               
                MethodInfo getPositionAtOffsetMethod = start.GetType().GetMethod("CreatePointer", new Type[] { start.GetType(), typeof(int) });
            }

            if (rank[currentIndex] >= embeddings.Length)
            {
                var i = rank[currentIndex] - embeddings.Length;
                var region = paths[i].Data.Bounds;
                // Create a semi-transparent rectangle to represent the highlight.
                System.Windows.Shapes.Rectangle highlight = new System.Windows.Shapes.Rectangle
                {
                    Width = region.Width,
                    Height = region.Height,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(128, 0, 0, 255)) // Semi-transparent blue
                };


                canvases[i].Children.Add(highlight);
                System.Windows.Controls.Canvas.SetLeft(highlight, System.Windows.Controls.Canvas.GetLeft(paths[i]));
                System.Windows.Controls.Canvas.SetTop(highlight, System.Windows.Controls.Canvas.GetTop(paths[i]));

                // Position the highlight rectangle at the specified coordinates.
                System.Windows.Controls.Canvas.SetLeft(highlight, region.Left);
                System.Windows.Controls.Canvas.SetTop(highlight, region.Top);

                imageHighligted = true;
                highlightRect = highlight;

                return;
            }

            imageHighligted = false;
            object highlighStartPointer = textPointersStart[rank[currentIndex]];
            object highlighEndPointer = textPointersEnd[rank[currentIndex]];

            // Get the TextRange constructor that takes two ITextPointers
            Type textRangeType = typeof(System.Windows.Documents.TextRange);

            // Clearing the selection in order to start search from the beginning of the document. I suspect there might be a better way of doing this.
            object segmentStart = textSegments[0].GetType().GetField("_start", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(textSegments[0]); // get segment start (one textsegment is always present)
            int currentOffset = (int)segmentStart.GetType().GetProperty("System.Windows.Documents.ITextPointer.Offset", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(segmentStart); // get offset of segment start
            segmentStart = segmentStart.GetType().GetMethod("CreatePointer", new Type[] { segmentStart.GetType(), typeof(int) }).Invoke(segmentStart, new object[] { segmentStart, -currentOffset }); // set the offset back to 0

            textSegments[0] = textSegments[0].GetType().GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[] { segmentStart.GetType(), segmentStart.GetType() }, null)
                                                       .Invoke(new object[] { segmentStart, segmentStart }); // create a new textsegment with resetted offset
            searchTextSegment = textSegments[0].GetType().GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[] { segmentStart.GetType(), segmentStart.GetType() }, null)
                                .Invoke(new object[] { highlighStartPointer, highlighEndPointer }); // create a new textsegment with resetted offset

            //allSegments.Add(searchTextSegment);
            
            for (int i = 1; i < textSegments.Count; i++)
            {
                textSegments.RemoveAt(i); // remove all other segments
            }
            // Get the FindTextBox control from the FindToolBar
            System.Windows.Controls.TextBox findTextBox = findToolbar.GetType().GetField("FindTextBox", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(_myfindToolbar) as System.Windows.Controls.TextBox;

            // Always search down
            _myfindToolbar.GetType().GetProperty("SearchUp").SetValue(_myfindToolbar, false);

            // Search and collect the find results
            int resultCount = 0;
          //  do
            {
                // invoke: DocumentViewerBase.Find(findToolBar)
                findResult = _miFind.Invoke(this, new object[] { _myfindToolbar }) as System.Windows.Documents.TextRange;

                /*if (findResult != null)
                {
                    // get the selected TextSegments of the search
                    textSegments = fiTextSegments.GetValue(selection) as IList; // List<System.Windows.Documents.TextSegment>
                    if (allSegments == null)
                        allSegments = textSegments; // first search find, set whole collection
                    else
                        allSegments.Add(textSegments[0]); // after first find, add to collection

                    resultCount++;
                }*/
            }
           // while (findResult != null && (MaxSearchResults == 0 || resultCount < MaxSearchResults)); // stop if no more results were found or limit is exceeded

/*            if (allSegments == null)
            {
                // alert the user that we did not find anything
                string searchText = _myfindToolbar.GetType().GetProperty("SearchText").GetValue(_myfindToolbar) as string;
                string messageString = string.Format("Searched the document. Cannot find '{0}'.", searchText);

                MessageBox.Show(messageString, "Find", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else*/
            {
                // Call the Clear method on the TextSelection instance
                this.Focus();

                textSegments.Clear(); // remove the first segment, which is the search text
                textSegments.Add(searchTextSegment);

                MethodInfo selectMethod = selection.GetType().GetMethod("SetActivePositions", BindingFlags.NonPublic | BindingFlags.Instance);
                selectMethod.Invoke(selection, new object[] { start, start });

                // set the textsegments field to the collected search results
                fiTextSegments.SetValue(selection, textSegments);

                // this marks the text. invoke: DocumentGrid.MakeSelectionVisible()
                _miMakeSelectionVisible.Invoke(_mydocumentScrollInfo, null);

            }

            // put the focus back on the findtoolbar textbox to search again. invoke: FindToolBar.GoToTextBox()
            _miGoToTextBox.Invoke(_myfindToolbar, null);

        }

        public async Task SetDocText(string text)
        {
            docText = text;

            object textEditor = this.GetType().GetProperty("TextEditor", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(this); // System.Windows.Documents.TextEditor

            // Get the TextContainer from the TextEditor
            FieldInfo piTextContainer = textEditor.GetType().GetPrivateFieldOfBase("_textContainer");
            object textContainer = piTextContainer.GetValue(textEditor);
            // Get the Start and End properties of the ITextContainer
            FieldInfo startProp = textContainer.GetType().GetPrivateFieldOfBase("_start");
            FieldInfo endProp = textContainer.GetType().GetPrivateFieldOfBase("_end");

            // Get the Start and End ITextPointers
            object start = startProp.GetValue(textContainer);
            object end = endProp.GetValue(textContainer);


            // Get the Start TextPointer
            object pointer = startProp.GetValue(textContainer);

            textPointersStart = new List<object>();
            textPointersEnd = new List<object>();
            actualSentences = new List<string>();
            bitmapSource = new List<BitmapSource>();
            paths = new List<System.Windows.Shapes.Path>();
            canvases = new List<System.Windows.Controls.Canvas>();

            int offset = 0;

            while (pointer != null)
            {
                // Get the GetPointerContext method of the TextPointer
                MethodInfo getPointerContextMethod = pointer.GetType().GetMethod("GetPointerContext");

                // Call the GetPointerContext method
                object context = getPointerContextMethod.Invoke(pointer, new object[] { pointer, System.Windows.Documents.LogicalDirection.Forward });

                if ((TextPointerContext)context == TextPointerContext.Text)
                {
                    // Get the GetTextInRun method of the TextPointer
                    MethodInfo getTextInRunMethod = pointer.GetType().GetMethod("System.Windows.Documents.ITextPointer.GetTextInRun", BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[] { typeof(System.Windows.Documents.LogicalDirection) }, null);

                    // Call the GetTextInRun method
                    string textRun = (string)getTextInRunMethod.Invoke(pointer, new object[] { System.Windows.Documents.LogicalDirection.Forward });

                    if (textRun.Length <= 2)
                    {
                        // Get the GetNextContextPosition method of the TextPointer
                        MethodInfo getNextContextPositionMethod1 = pointer.GetType().GetMethod("System.Windows.Documents.ITextPointer.GetNextContextPosition", BindingFlags.NonPublic | BindingFlags.Instance);

                        // Call the GetNextContextPosition method
                        pointer = getNextContextPositionMethod1.Invoke(pointer, new object[] { System.Windows.Documents.LogicalDirection.Forward });
                        continue; 
                    }
                    var textList = textRun.Split(". ").ToList();

                    offset = 0;
                    int prev_offset = 0;


                    for (int i = 0; i < textList.Count(); i++)
                    {
                        if (textList[i] == "")
                        { continue; }
                        else if (textList[i].Length <= 2)
                        {
                            offset += textList[i].Length + 2;

                            if (i == textList.Count() - 1)
                            {
                                offset -= 2;
                            }

                            continue;
                        }
                    
                        int sentenceLength = textList[i].Length + 2;

                        if (i == textList.Count() - 1)
                        {
                            sentenceLength -= 2;
                        }
                        prev_offset = offset;
                        offset = offset + sentenceLength;

                        actualSentences.Add(textList[i]);


                        // Get the GetPositionAtOffset method of the ITextPointer
                        MethodInfo getPositionAtOffsetMethod = pointer.GetType().GetMethod("CreatePointer", new Type[] { pointer.GetType(), typeof(int) });

                        // Create an ITextPointer for the start of the searchText
                        object searchTextStart = getPositionAtOffsetMethod.Invoke(pointer, new object[] { pointer, prev_offset });

                        // Create an ITextPointer for the end of the searchText
                        object searchTextEnd = getPositionAtOffsetMethod.Invoke(pointer, new object[] { pointer, offset });

                        textPointersStart.Add(searchTextStart);
                        textPointersEnd.Add(searchTextEnd); 
                    }

                    offset += textRun.Length;
                }
                // Get the GetNextContextPosition method of the TextPointer
                MethodInfo getNextContextPositionMethod = pointer.GetType().GetMethod("System.Windows.Documents.ITextPointer.GetNextContextPosition", BindingFlags.NonPublic | BindingFlags.Instance);

                // Call the GetNextContextPosition method
                pointer = getNextContextPositionMethod.Invoke(pointer, new object[] { System.Windows.Documents.LogicalDirection.Forward });
            }

            FixedDocumentSequence fds = this.Document as FixedDocumentSequence;
List<System.Windows.Media.Imaging.BitmapSource> bitmaps = new List<System.Windows.Media.Imaging.BitmapSource>();
          //  bitmapSource = new List<BitmapSource>();

            foreach (DocumentReference docRef in fds.References)
            {
                FixedDocument doc = docRef.GetDocument(false);
                foreach (PageContent pageContent in doc.Pages)
                {
                    FixedPage page = (FixedPage)pageContent.GetPageRoot(false);
                    var children = page.Children;
                    foreach (UIElement element in page.Children)
                    {
                        if (element is System.Windows.Controls.Canvas canvas)
                        {
                            TraverseCanvas(canvas);
                        }
                    }
                }
            }

            /*// Get the FlowDocument
            object doc1 = this.Document;
            FlowDocument flowDocument = this.Document as FlowDocument;

            // Get the FlowDocument type
            Type flowDocumentType = typeof(FlowDocument);

            // Get the Blocks property
            PropertyInfo blocksProperty = flowDocumentType.GetProperty("Blocks");

            // Get the BlockCollection
            object blockCollection = blocksProperty.GetValue(flowDocument); 

            System.Windows.Documents.BlockCollection blocks = blockCollection as System.Windows.Documents.BlockCollection;
            // Traverse the BlockCollection
            foreach (System.Windows.Documents.Block block in blocks)
            {
                if (block is System.Windows.Documents.Paragraph)
                {
                    System.Windows.Documents.Paragraph paragraph = block as System.Windows.Documents.Paragraph;
                    foreach (System.Windows.Documents.Inline inline in paragraph.Inlines)
                    {
                        if (inline is System.Windows.Documents.InlineUIContainer)
                        {
                            System.Windows.Documents.InlineUIContainer inlineUIContainer = inline as System.Windows.Documents.InlineUIContainer;
                            if (inlineUIContainer.Child is System.Windows.Controls.Image)
                            {
                                System.Windows.Controls.Image image = inlineUIContainer.Child as System.Windows.Controls.Image;
                                System.Windows.Media.Imaging.BitmapSource bitmapSource = image.Source as System.Windows.Media.Imaging.BitmapSource;

                                // Now you have the BitmapSource of the image
                                // You can convert it to a Bitmap if you need to

                                // To get the TextPointer to the image
                                //TextPointer textPointer = inlineUIContainer.ContentStart;

                                // Now you have the TextPointer to the image
                            }
                        }
                    }
                }
            }*/

            await InitializeEmbeddings();
        }

        void TraverseCanvas(System.Windows.Controls.Canvas canvas)
        {
            foreach (UIElement child in canvas.Children)
            {
                if (child is System.Windows.Controls.Canvas childCanvas)
                {
                    // Recursively traverse child canvases
                    TraverseCanvas(childCanvas);
                  //  return;
                }
                else if (child is System.Windows.Shapes.Path path)
                {
                    var bounds = path.Data.GetRenderBounds(null);
                    path.Measure(bounds.Size);
                    path.Arrange(bounds);
                    if (bitmapSource.Count != 0)
                    {
                        SolidColorBrush sr = new SolidColorBrush(System.Windows.Media.Color.FromArgb(128, 0, 0, 255)); // Semi-transparent yellow
                        sr.Opacity = 0.8;
                        // path.Fill = sr;
                        
                        Rectangle highlight = new Rectangle
                        {
                            Width = bounds.Width,
                            Height = bounds.Height,
                            Fill = new SolidColorBrush(Color.FromArgb(128, 0, 0, 255)) // Semi-transparent blue
                        };

                    }
                    var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(
                            (int)bounds.Width, (int)bounds.Height, 96, 96, PixelFormats.Pbgra32);
                        bitmap.Render(path);
                    // bitmap.
                    BitmapSource bitSource = System.Windows.Media.Imaging.BitmapFrame.Create(bitmap);
                    bitmapSource.Add(bitSource);
                    paths.Add(path);

                    canvases.Add(canvas);
                }
            }
        }

        public List<BitmapSource> GetBitmapSource()
        {
            return bitmapSource;
        }

        public List<System.Windows.Shapes.Path> GetPaths()
        {
            return paths;
        }
        
        private async Task InitializeEmbeddings()
        {
            if (!SemanticTextEmbeddingsCreator.IsAvailable())
            {
                SemanticTextEmbeddingsCreator.MakeAvailableAsync().AsTask().Wait();
            }

            if (embedCreator == null) { embedCreator = await ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Text); }
            currentIndex = 0;

            embeddings = new EmbeddingVector[actualSentences.Count];
            for (int i = 0; i < actualSentences.Count; i++)
            {

                embeddings[i] = await embedCreator.CreateVectorForTextAsync(actualSentences[i]);
            }

            imageEmbeddings = new EmbeddingVector[bitmapSource.Count];


            for (int i = 0; i < bitmapSource.Count; i++)
            {
                BitmapSource bitSource = bitmapSource[i]; // Your BitmapSource

                // Get pixels as an array of bytes
                var stride = bitSource.PixelWidth * bitSource.Format.BitsPerPixel / 8;
                var bytes = new byte[stride * bitSource.PixelHeight];
                bitSource.CopyPixels(bytes, stride, 0);
                var buffer = bytes.AsBuffer();

                // Create an ImageBuffer
                ImageBuffer imageBuffer = new ImageBuffer(buffer, Microsoft.Windows.Imaging.PixelFormat.Bgra32, (uint) bitSource.PixelWidth, (uint) bitSource.PixelHeight);
                //imageBuffer.CopyFromBuffer(bytes);

                if (imageEmbedCreator == null) { imageEmbedCreator = await ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Image); }

                imageEmbeddings[i] = await imageEmbedCreator.CreateVectorForImageAsync(imageBuffer);
            }
        }

        public static Microsoft.Windows.Imaging.PixelFormat GetPixelFormatFromBitmapPixelFormat(BitmapPixelFormat bitmapPixelFormat)
        {
            switch (bitmapPixelFormat)
            {
                case BitmapPixelFormat.Bgra8:
                    return Microsoft.Windows.Imaging.PixelFormat.Bgra32;

                case BitmapPixelFormat.Rgba8:
                    return Microsoft.Windows.Imaging.PixelFormat.Rgba32;

                default:
                    throw new ArgumentException();
            }
        }


        public static float CheckOverflow(double x)
        {
            if (x >= double.MaxValue)
            {
                throw new OverflowException("operation caused overflow");
            }
            return (float)x;
        }

        public static float DotProduct(float[] a, float[] b)
        {
            float result = 0.0f;
            for (int i = 0; i < a.Length; i++)
            {
                result = CheckOverflow(result + CheckOverflow(a[i] * b[i]));
            }
            return result;
        }

        public static float Magnitude(float[] v)
        {
            float result = 0.0f;
            for (int i = 0; i < v.Length; i++)
            {
                result = CheckOverflow(result + CheckOverflow(v[i] * v[i]));
            }
            return (float)Math.Sqrt(result);
        }

        public static float CosineSimilarity(EmbeddingVector vector1, EmbeddingVector vector2)
        {
            if (vector1.Count != vector2.Count)
            {
                throw new ArgumentException("Vector lengths must be equal");
            }

            float[] vec1 = new float[vector1.Count];
            vector1.GetValues(vec1);

            float[] vec2 = new float[vector2.Count];
            vector2.GetValues(vec2);

            float dotProduct = 0;
            float norm1 = 0;
            float norm2 = 0;
            for (int i = 0; i < vec1.Length; i++)
            {
                dotProduct += vec1[i] * vec2[i];
                norm1 += vec1[i] * vec1[i];
                norm2 += vec2[i] * vec2[i];
            }
            float score = dotProduct / (MathF.Sqrt(norm1) * MathF.Sqrt(norm2));

            return score;
        }

        public static float CosineSimilarity2(EmbeddingVector v1, EmbeddingVector v2)
        {
            if (v1.Count != v2.Count)
            {
                throw new ArgumentException("Vectors must have the same length.");
            }

            int size = (int)(v1.Count);

            float[] raw1 = new float[size];
            float[] raw2 = new float[size];
            v1.GetValues(raw1);
            v2.GetValues(raw2);
/*            float m1 = Magnitude(raw1);
            float m2 = Magnitude(raw2);
                                    var normalizedList1 = raw1.Select(o => o / m1).ToArray();


                                    var normalizedList2 = raw2.Select(o => o / m2).ToArray();
*/            


            /*// Vectors should already be normalized.
            if (Math.Abs(m1 - m2) > 0.4f || Math.Abs(m1 - 1.0f) > 0.4f)
            {
                throw new InvalidOperationException("Vectors are not normalized.");
            }*/

            return DotProduct(raw1, raw2);
        }


    }

    public static class ReflectionExtensions
    {
        /// <summary>
        /// Gets private field of base class. Normally, they are not directly accessible in a GetField call.
        /// </summary>
        public static FieldInfo GetPrivateFieldOfBase(this Type type, string fieldName)
        {
            BindingFlags bindingFlags = BindingFlags.Instance | BindingFlags.NonPublic;

            // Declare variables
            FieldInfo fieldInfo = null;

            // Search as long as there is a type
            while (type != null)
            {
                // Use reflection
                fieldInfo = type.GetField(fieldName, bindingFlags);

                // Yes, do we have a field?
                if (fieldInfo != null) break;

                // Get base class
                type = type.BaseType;
            }

            // Return result
            return fieldInfo;
        }
    }

    /// <summary>
    /// http://www.codeproject.com/Articles/103542/Removing-Event-Handlers-using-Reflection
    /// </summary>
    public static class ReflectionHelper
    {
        static Dictionary<Type, List<FieldInfo>> dicEventFieldInfos = new Dictionary<Type, List<FieldInfo>>();

        static BindingFlags AllBindings
        {
            get { return BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static; }
        }

        static List<FieldInfo> GetTypeEventFields(Type t)
        {
            if (dicEventFieldInfos.ContainsKey(t))
                return dicEventFieldInfos[t];

            List<FieldInfo> lst = new List<FieldInfo>();
            BuildEventFields(t, lst);
            dicEventFieldInfos.Add(t, lst);
            return lst;
        }

        static void BuildEventFields(Type t, List<FieldInfo> lst)
        {
            //BindingFlags.FlattenHierarchy only works on protected & public, doesn't work because fields are private
            // Uses .GetEvents and then uses .DeclaringType to get the correct ancestor type so that we can get the FieldInfo.
            foreach (EventInfo ei in t.GetEvents(AllBindings))
            {
                Type dt = ei.DeclaringType;
                FieldInfo fi = dt.GetField(ei.Name, AllBindings);
                if (fi != null)
                    lst.Add(fi);
            }
        }

        static EventHandlerList GetStaticEventHandlerList(Type t, object obj)
        {
            MethodInfo mi = t.GetMethod("get_Events", AllBindings);
            return (EventHandlerList)mi.Invoke(obj, new object[] { });
        }

        public static void RemoveAllEventHandlers(object obj) { RemoveEventHandler(obj, ""); }

        public static void RemoveEventHandler(object obj, string EventName)
        {
            if (obj == null)
                return;

            Type t = obj.GetType();
            List<FieldInfo> event_fields = GetTypeEventFields(t);
            EventHandlerList static_event_handlers = null;

            foreach (FieldInfo fi in event_fields)
            {
                if (EventName != "" && string.Compare(EventName, fi.Name, true) != 0)
                    continue;

                // STATIC Events have to be treated differently from INSTANCE Events...
                if (fi.IsStatic)
                {
                    if (static_event_handlers == null)
                        static_event_handlers = GetStaticEventHandlerList(t, obj);

                    object idx = fi.GetValue(obj);
                    Delegate eh = static_event_handlers[idx];
                    if (eh == null)
                        continue;

                    Delegate[] dels = eh.GetInvocationList();
                    if (dels == null)
                        continue;

                    EventInfo ei = t.GetEvent(fi.Name, AllBindings);
                    foreach (Delegate del in dels)
                        ei.RemoveEventHandler(obj, del);
                }
                else
                {
                    EventInfo ei = t.GetEvent(fi.Name, AllBindings);
                    if (ei != null)
                    {
                        object val = fi.GetValue(obj);
                        Delegate mdel = (val as Delegate);
                        if (mdel != null)
                        {
                            foreach (Delegate del in mdel.GetInvocationList())
                                ei.RemoveEventHandler(obj, del);
                        }
                    }
                }
            }
        }
    }
}
