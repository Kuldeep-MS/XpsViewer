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
using Microsoft.Windows.Vision;

using static System.Net.Mime.MediaTypeNames;
using Microsoft.Windows.Imaging;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Graphics.Imaging;
using System.Windows.Input;

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
        public List<EmbeddingVector> imageTextEmbeddings;
        public List<int> imageIndexesForText;
        public int[] indexes;
        public float[] scores;
        public int[] rank;
        public int currentIndex = 0;
        public int[] rankImages;
        public int currentIndexImages = 0;
        public float[] scoresImage;
        public int[] uniqueIndexes;
        public List<double> offsetList = new List<double>();
        public bool updated = false;
        ImageSearchEmbeddingsCreator embedCreator;
        ImageSearchEmbeddingsCreator imageEmbedCreator;
        String currentSearchString;
        String currentImageSearchString;

        IList textPointersStart;
        IList textPointersEnd;
        object searchTextSegment;

        public bool IsTextSearch = true;
        bool imageHighligted = false;
        System.Windows.Shapes.Rectangle highlightRect;
        private ToolBar _myfindToolbar; // MS.Internal.Documents.FindToolBar
        private object _mydocumentScrollInfo; // MS.Internal.Documents.DocumentGrid
        private System.Windows.Controls.ScrollViewer scrollViewer;
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
            //this.MouseLeftButtonDown += DocumentViewer_MouseLeftButtonDown;

            if (IsMultiSearchEnabled)
            {
                // get some private fields from the base class DocumentViewer
                _myfindToolbar = this.GetType().GetPrivateFieldOfBase("_findToolbar").GetValue(this) as ToolBar;
                _mydocumentScrollInfo = this.GetType().GetPrivateFieldOfBase("_documentScrollInfo").GetValue(this);
                scrollViewer = this.GetType().GetPrivateFieldOfBase("_scrollViewer").GetValue(this) as System.Windows.Controls.ScrollViewer;//, BindingFlags.NonPublic | BindingFlags.Instance);


                // Increase the size of the toolbar
                _myfindToolbar.Height = 40;
                _myfindToolbar.Width = 3000;
                Thickness margin = _myfindToolbar.Margin;
                margin.Left = 50;
                _myfindToolbar.Margin = margin;


                // Define your new icons
                System.Windows.Controls.Canvas newPreviousIcon = new System.Windows.Controls.Canvas(); // Customize this with your new icon
                System.Windows.Controls.Canvas newNextIcon = new System.Windows.Controls.Canvas(); // Customize this with your new icon

                // Replace the icons in the toolbar's resources
                // The keys "FindPreviousContent" and "FindNextContent" are based on your XAML context
                _myfindToolbar.Resources["FindPreviousContent"] = newPreviousIcon;
                _myfindToolbar.Resources["FindNextContent"] = newNextIcon;

                System.Windows.Controls.Button findNextButton = _myfindToolbar.FindName("FindNextButton") as System.Windows.Controls.Button;
                System.Windows.Controls.Button findPreviousButton = _myfindToolbar.FindName("FindPreviousButton") as System.Windows.Controls.Button;

                EventInfo clickEvent = typeof(System.Windows.Controls.Button).GetEvent("Click");
                // Remove the existing Click event handlers from the buttons
                ReflectionHelper.RemoveEventHandler(findNextButton, clickEvent.Name);
                ReflectionHelper.RemoveEventHandler(findPreviousButton, clickEvent.Name);

                // Add your own Click event handlers to the buttons
                clickEvent.AddEventHandler(findNextButton, new RoutedEventHandler(OnFindNextClick));
                clickEvent.AddEventHandler(findPreviousButton, new RoutedEventHandler(OnFindPreviousClick));

                // replace button click handler of find toolbar
                EventInfo evt = _myfindToolbar.GetType().GetEvent("FindClicked");
                ReflectionHelper.RemoveEventHandler(_myfindToolbar, evt.Name); // remove existing handler
                evt.AddEventHandler(_myfindToolbar, new EventHandler(OnFindInvoked)); // attach own handler


                // get some methods that will need to be invoked
                _miFind = this.GetType().GetMethod("Find", BindingFlags.NonPublic | BindingFlags.Instance);
                _miGoToTextBox = _myfindToolbar.GetType().GetMethod("GoToTextBox");
                _miMakeSelectionVisible = _mydocumentScrollInfo.GetType().GetMethod("MakeSelectionVisible");
            }
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            if (highlightRect != null)
            {
                canvases[uniqueIndexes[currentIndexImages]].Children.Remove(highlightRect);
                highlightRect = null;
            }

        }

        public void MouseLeftButtonDown2()
        {
            if (highlightRect != null)
            {
                canvases[uniqueIndexes[currentIndexImages]].Children.Remove(highlightRect);
                highlightRect = null;
            }

        }

        private void PerformTextSearch(bool forward)
        {
            IList allSegments = null; // collection of text segments
            System.Windows.Documents.TextRange findResult = null; // could also use object, does not need type

            // Drill down to the list of selected text segments: DocumentViewer.TextEditor.Selection.TextSegments
            object textEditor = this.GetType().GetProperty("TextEditor", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(this); // System.Windows.Documents.TextEditor
            object selection = textEditor.GetType().GetProperty("Selection", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(textEditor); // System.Windows.Documents.TextSelection
            FieldInfo fiTextSegments = selection.GetType().GetPrivateFieldOfBase("_textSegments");
            IList textSegments = fiTextSegments.GetValue(selection) as IList; // List<System.Windows.Documents.TextSegment>


            // Get the SearchText property
            PropertyInfo piSearchText = _myfindToolbar.GetType().GetProperty("SearchText", BindingFlags.Public | BindingFlags.Instance);

            // Get the searchable text
            string findSearchText = (string)piSearchText.GetValue(_myfindToolbar);

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

            

            if (currentSearchString == findSearchText && currentIndex >= 0)
            {
                if (currentIndex > 0 && forward == false)
                {
                    currentIndex--;
                }
                else if (currentIndex == 0 && forward == false)
                { }
                else
                {
                    currentIndex++;
                }
            }
            else
            {
                currentSearchString = findSearchText;
                currentIndex = 0;

                var inputStringEmbed = embedCreator.CreateVectorForText(currentSearchString);
                //var op = ImageSearchEmbeddingsCreator.MakeAvailableAsync();

                var model8 = ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Text).AsTask().Result;
                var op = ImageSearchEmbeddingsCreator.MakeAvailableAsync();
                //rank = embedCreator.CalculateRanking();

                scores = new float[embeddings.Length];
                rank = new int[embeddings.Length];

                for (int i = 0; i < embeddings.Length; i++)
                {
                    if (embeddings[i] == null)
                    {
                        scores[i] = 0;
                        continue;
                    }
                    var score = CosineSimilarity(embeddings[i], inputStringEmbed);
                    scores[i] = score;
                }


                var indexedFloats = scores.Select((value, index) => new { Value = value, Index = index })
                  .ToArray();

                // Sort the indexed floats by value in descending order
                Array.Sort(indexedFloats, (a, b) => b.Value.CompareTo(a.Value));

                // Extract the top k indices
                rank = indexedFloats.Select(item => item.Index).ToArray();

                int offset = 0;


                MethodInfo getPositionAtOffsetMethod = start.GetType().GetMethod("CreatePointer", new Type[] { start.GetType(), typeof(int) });
            }

            if (currentIndex >= rank.Length || scores[rank[currentIndex]] < 0.80)
            { return; }
            //   imageHighligted = false;
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
            System.Windows.Controls.TextBox findTextBox = _myfindToolbar.GetType().GetField("FindTextBox", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(_myfindToolbar) as System.Windows.Controls.TextBox;

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
        private void OnFindNextClick(object sender, RoutedEventArgs e)
        {
            if (IsTextSearch)
            {
                PerformTextSearch(true);
            }
            else
            {
                PerformImageSearch(true);
            }

            
        }

        private void PerformImageSearch(bool forward)
        {
            // Get the SearchText property
            PropertyInfo piSearchText = _myfindToolbar.GetType().GetProperty("SearchText", BindingFlags.Public | BindingFlags.Instance);

            // Get the searchable text
            string findSearchText = (string)piSearchText.GetValue(_myfindToolbar);

            if (highlightRect != null)
            {
                if (currentIndexImages < uniqueIndexes.Length)
                {
                    canvases[uniqueIndexes[currentIndexImages]].Children.Remove(highlightRect);
                }
            }

            if (currentImageSearchString == findSearchText && currentIndexImages >= 0)
            {
                if (currentIndexImages > 0 && forward == false)
                {
                    currentIndexImages = currentIndexImages - 1;
                }
                else if (currentIndexImages == 0 && forward == false)
                { return; }
                else if (currentIndexImages == uniqueIndexes.Length - 1 && forward == true)
                { return; }
                else
                {
                    currentIndexImages++;
                }
            }
            else
            {
                currentImageSearchString = findSearchText;
                currentIndexImages = 0;
                var model8 = ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Text).AsTask().Result;

                var embedImageText = model8.CreateVectorForText(currentImageSearchString);

                scoresImage = new float[imageEmbeddings.Length];

                for (int i = 0; i < imageEmbeddings.Length; i++)
                {
                    var score = CosineSimilarity2(imageEmbeddings[i], embedImageText);
                    scoresImage[i] = score;
                }

                var scoresImageText = new float[imageTextEmbeddings.Count];
                // calculate score for imageTextEmbeddings
                for (int i = 0; i < imageTextEmbeddings.Count; i++)
                {
                    var score = CosineSimilarity2(imageTextEmbeddings[i], embedImageText);
                    float scaling = 0.3f;
                    scoresImageText[i] = score * scaling;
                }

                // ramk the scores by index for imageTextEmbeddings
                var indexedFloatsImageText = scoresImageText.Select((value, index) => new { Value = value, Index = index })
                  .ToArray();
                Array.Sort(indexedFloatsImageText, (a, b) => b.Value.CompareTo(a.Value));


                var rankImageText = indexedFloatsImageText.Select(item => imageIndexesForText[item.Index]).ToArray();

                var indexedFloats = scoresImage.Select((value, index) => new { Value = value, Index = index }).ToArray();

                // Sort the indexed floats by value in descending order
                Array.Sort(indexedFloats, (a, b) => b.Value.CompareTo(a.Value));

                // Extract the top k indices
                rankImages = indexedFloats.Select(item => item.Index).ToArray();

                // for rankImages and rankImageText, create a new array wit both where the indexes are sorted by their corresponding scoresImage and scoresImageText
                // Create a list of (index, score) pairs for images
                var imageScoreList = scoresImage.Select((score, index) => new { Index = index, Score = score });

                // Create a list of (index, score) pairs for image text
                var imageTextScoreList = scoresImageText.Select((score, index) => new { Index = imageIndexesForText[index], Score = score });

                // Concatenate the two lists
                var combinedList = imageScoreList.Concat(imageTextScoreList);

                // Sort by score and select the indexes
                var sortedIndexes = combinedList.OrderByDescending(x => x.Score).Select(x => x.Index).ToArray();

                // remove the repeated occurring indexes
                uniqueIndexes = sortedIndexes.Distinct().ToArray();

            }

            if (currentIndexImages >= uniqueIndexes.Length)// || scoresImage[rankImages[currentIndexImages]] < 0.31)
            {
                return;
            }
            var j = uniqueIndexes[currentIndexImages];

            var region = paths[j].Data.Bounds;
            // Create a semi-transparent rectangle to represent the highlight.
            System.Windows.Shapes.Rectangle highlight = new System.Windows.Shapes.Rectangle
            {
                Width = region.Width,
                Height = region.Height,
                Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(128, 0, 0, 255)) // Semi-transparent blue
            };
            scrollViewer.ScrollToVerticalOffset(offsetList[j]);
            canvases[j].Children.Add(highlight);
            System.Windows.Controls.Canvas.SetLeft(highlight, System.Windows.Controls.Canvas.GetLeft(paths[j]));
            System.Windows.Controls.Canvas.SetTop(highlight, System.Windows.Controls.Canvas.GetTop(paths[j]));

            // Position the highlight rectangle at the specified coordinates.
            System.Windows.Controls.Canvas.SetLeft(highlight, region.Left);
            System.Windows.Controls.Canvas.SetTop(highlight, region.Top);

            imageHighligted = true;
            highlightRect = highlight;
        }

        private void OnFindPreviousClick(object sender, RoutedEventArgs e)
        {
            if (IsTextSearch)
            {
                PerformTextSearch(false);
            }
            else
            {
                PerformImageSearch(false);
            }
        }

        private void OnFindInvoked(object sender, EventArgs e)
        {
            /*if (IsTextSearch)
            {
                PerformTextSearch(true);
            }
            else
            {
                PerformImageSearch(true);
            }*/
        }
            
        public async Task SetDocText(string text)
        {
            docText = text;
            currentIndex = -1;
            currentIndexImages = -1;
            highlightRect = null;
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
                }
                else if (child is System.Windows.Shapes.Path path)
                {
                    // Check if the Fill property of the Path is an ImageBrush
                    if (path.Fill is ImageBrush imageBrush)
                    {
                        // Extract the ImageSource from the ImageBrush
                        ImageSource imageSource = imageBrush.ImageSource;

                        // Check if the ImageSource is a BitmapImage
                        if (imageSource is BitmapSource bitSource)
                        {
                            // Get the file path from the UriSource of the BitmapImage

                            // Add the BitmapSource to your collection
                            bitmapSource.Add(bitSource);
                        }


                        var bounds = path.Data.GetRenderBounds(null);
                        path.Measure(bounds.Size);
                        path.Arrange(bounds);
                        if (bitmapSource.Count != 0)
                        {

                            Rectangle highlight = new Rectangle
                            {
                                Width = bounds.Width,
                                Height = bounds.Height,
                                Fill = new SolidColorBrush(Color.FromArgb(128, 0, 0, 255)) // Semi-transparent blue
                            };
                        }

                        try
                        {
                            /*                            var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(
                                                            (int)bounds.Width, (int)bounds.Height, 96, 96, PixelFormats.Pbgra32);
                                                        bitmap.Render(path);
                                                        BitmapSource bitSource = System.Windows.Media.Imaging.BitmapFrame.Create(bitmap);
                                                        bitmapSource.Add(bitSource);

                            */
                                                       //get the position of path object
                                                      //  var position = System.Windows.Controls.Canvas.GetLeft(path);
                            // Assuming 'this' is an instance of DocumentViewer

                            // Assuming 'path' is the Path element you want to scroll to
                            // and 'canvas' is the Canvas containing the Path
                            scrollViewer = this.GetType().GetPrivateFieldOfBase("_scrollViewer").GetValue(this) as System.Windows.Controls.ScrollViewer;
                            GeneralTransform childTransform = path.TransformToAncestor(canvas);
                            Rect rectangle = childTransform.TransformBounds(new Rect(new Point(0, 0), path.RenderSize));

                            FixedDocumentSequence fds = this.Document as FixedDocumentSequence;
                            List<System.Windows.Media.Imaging.BitmapSource> bitmaps = new List<System.Windows.Media.Imaging.BitmapSource>();

                            foreach (DocumentReference docRef in fds.References)
                            {
                                FixedDocument doc = docRef.GetDocument(false);
                                int pageNumber = 0;
                                bool found = false;
                                foreach (PageContent pageContent in doc.Pages)
                                {
                                    pageNumber++;

                                    FixedPage page = (FixedPage)pageContent.GetPageRoot(false);
                                    try
                                    {
                                        path.TransformToAncestor(page);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine(ex.Message);
                                        continue;

                                    }
                                    if (path.TransformToAncestor(page) != null)
                                    {
                                        // Get the position of the Canvas in the FixedPage
                                        GeneralTransform transform = path.TransformToAncestor(page);
                                        Point position = transform.Transform(new Point(0, 0));

                                        double offset = (pageNumber - 1) * (page.Height + page.Margin.Top + page.Margin.Bottom) + position.Y ;
                                        // Scroll to the position of the Canvas
                                        offsetList.Add(offset);
                                        found = true;
                                        scrollViewer.ScrollToVerticalOffset(offset);
                                        break;
                                       // scrollViewer.ScrollToHorizontalOffset(position.X);
                                    }
                                }

                                if (found)
                                {
                                    break;
                                }
                            }

                            paths.Add(path);

                            canvases.Add(canvas);
                        }
                        catch (Exception ex)
                        {
                            // Handle the exception (e.g., log it)
                            Console.WriteLine(ex.Message);
                            continue; // Continue to the next iteration of the loop
                        }
                    }    

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
            imageEmbeddings = new EmbeddingVector[bitmapSource.Count];

            // create instance of imageembeddingcreator for type text
            var textImageEmbedCreator = await ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Text);
            // initialize imagetextembeddings and imageindexesfortext
            imageTextEmbeddings = new List<EmbeddingVector>();
            imageIndexesForText = new List<int>();
            
            for (int i = 0; i < bitmapSource.Count; i++)
            {
                BitmapSource bitSource = bitmapSource[i]; // Your BitmapSource
                var stride = bitSource.PixelWidth * bitSource.Format.BitsPerPixel / 8;
                var bytes = new byte[stride * bitSource.PixelHeight];
                bitSource.CopyPixels(bytes, stride, 0);
                var buffer = bytes.AsBuffer();
                var format = bitSource.Format;
                // Create an ImageBuffer
                ImageBuffer imageBuffer = new ImageBuffer(buffer, Microsoft.Windows.Imaging.PixelFormat.Bgra32, (uint) bitSource.PixelWidth, (uint) bitSource.PixelHeight);
                //imageBuffer.CopyFromBuffer(bytes);

                Microsoft.Windows.Vision.TextRecognizer textRecognizer = await TextRecognizer.CreateAsync();
                var textRecognized = await textRecognizer.RecognizeTextFromImageAsync(imageBuffer, null);
                string text = "";

                if (textRecognized.Lines != null)
                {
                    // Loop through the recognized text lines
                    foreach (var line in textRecognized.Lines)
                    {
                        // Extract the text from the line
                        text = line.Text;
                        try
                        {
                            var textEmbedding = await textImageEmbedCreator.CreateVectorForTextAsync(text);
                            //update imageTextEmbeddings and imageIndexesForText
                            imageTextEmbeddings.Add(textEmbedding);
                            imageIndexesForText.Add(i);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
                if (imageEmbedCreator == null) 
                { 
                    imageEmbedCreator = await ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Image);
                }

                imageEmbeddings[i] = await imageEmbedCreator.CreateVectorForImageAsync(imageBuffer);
            }

            if (!SemanticTextEmbeddingsCreator.IsAvailable())
            {
                SemanticTextEmbeddingsCreator.MakeAvailableAsync().AsTask().Wait();
            }

            if (embedCreator == null) { embedCreator = await ImageSearchEmbeddingsCreator.CreateAsync(ImageSearchEmbeddingsType.Text); }
            currentIndex = 0;

            embeddings = new EmbeddingVector[actualSentences.Count];
            for (int i = 0; i < actualSentences.Count; i++)
            {
                try
                {
                    embeddings[i] = await embedCreator.CreateVectorForTextAsync(actualSentences[i]);
                }
                catch (Exception ex)
                {
                    embeddings[i] = null;
                }
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
