using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Reactive.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DynamicData;
using ReactiveUI;

namespace DocxTestBuilder.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        public ReactiveCommand<Unit,Unit> OpenFile { get; }
        public ReactiveCommand<Unit,Unit> CopyValue { get; }
        public ObservableCollection<DocumentTreeNode> DocumentTree { get; }
        public ObservableCollection<string> NodeProperties { get; }

        public DocumentTreeNode CurrentNode
        {
            get => _currentNode;
            set => this.RaiseAndSetIfChanged(ref _currentNode, value);
        }

        public string CurrentPropertyName
        {
            get => _currentPropertyName;
            set => this.RaiseAndSetIfChanged(ref _currentPropertyName, value);
        }
        
        public string CurrentPropertyValue
        {
            get => _currentPropertyValue;
            set => this.RaiseAndSetIfChanged(ref _currentPropertyValue, value);
        }
        
        private DocumentTreeNode _currentNode;
        private string _currentPropertyName;
        private string _currentPropertyValue;

        public MainWindowViewModel()
        {
            NodeProperties = new ObservableCollection<string>();
            DocumentTree = new ObservableCollection<DocumentTreeNode>();
            OpenFile = ReactiveCommand.CreateFromTask(async () =>
            {
                var openFileDialog = new OpenFileDialog();
                openFileDialog.Filters = new List<FileDialogFilter>()
                    { new() { Name = "Word Documents", Extensions = { "docx" } } };
                if (Avalonia.Application.Current.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                {
                    var result = await openFileDialog.ShowAsync(desktop.MainWindow) ?? Array.Empty<string>();
                    if (result.Length > 0)
                    {
                        using var doc = WordprocessingDocument.Open(result[0], true);
                        DocumentTree.Clear();
                        foreach (var element in CreateTreeNodes(doc))
                        {
                            DocumentTree.Add(element);
                        }
                    }
                }
            });
            CopyValue = ReactiveCommand.CreateFromTask(async () =>
            {
                var data = new DataObject();
                var path = CurrentNode.Path;
                if (CurrentPropertyName != null)
                {
                    path += $".{CurrentPropertyName}";
                }
                data.Set("data",  $"{path}\t{CurrentPropertyValue}");
                await Application.Current.Clipboard.SetTextAsync($"{path}\t{CurrentPropertyValue}");
            });
            this.WhenAnyValue(x => x.CurrentNode).WhereNotNull().Subscribe(item =>
            {
                NodeProperties.Clear();
                CurrentPropertyName = null;
                NodeProperties.AddRange(GetProperties(item));
            });
            this.WhenAnyValue(x => x.CurrentPropertyName).Subscribe(item =>
            {
                if (item == null)
                {
                    CurrentPropertyValue = null;
                    return;
                }
                if (item == "InnerText")
                {
                    CurrentPropertyValue = CurrentNode.Tag.InnerText;
                }
                CurrentPropertyValue = CurrentNode.Tag.GetAttributes().FirstOrDefault(i => i.LocalName == item).Value;
            });
        }
        private List<DocumentTreeNode> CreateTreeNodes(WordprocessingDocument doc)
        {
            var output = new List<DocumentTreeNode>();
            if (doc.MainDocumentPart != null)
            {
                var documentNode = new DocumentTreeNode("document");
                foreach (var node in doc.MainDocumentPart.Document.ChildElements)
                {
                    documentNode.Children.Add(ConvertToTreeNode(node, $"document"));
                }
                output.Add(documentNode);
            }
            var footNoteNode = new DocumentTreeNode("footnotes");
            foreach (var node in doc.MainDocumentPart.FootnotesPart.Footnotes.ChildElements)
            {
                footNoteNode.Children.Add(ConvertToTreeNode(node, $"footnotes"));
            }
            output.Add(footNoteNode);
            var settingsNode = new DocumentTreeNode("settings");
            foreach (var node in doc.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements)
            {
                settingsNode.Children.Add(ConvertToTreeNode(node, $"settings"));
            }
            output.Add(settingsNode);
            /*
            var styleNode = new DocumentTreeNode("styles");
            foreach (var node in doc.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements)
            {
                settingsNode.Children.Add(ConvertToTreeNode(node, $"styles"));
            }
            output.Add(styleNode);
            */
            return output;
        }

        private DocumentTreeNode ConvertToTreeNode(OpenXmlElement input, string path, int occurenceNumber = 0)
        {
            var nodeCount = new Dictionary<string, int>();
            var output = new DocumentTreeNode(input, occurenceNumber,path);
            foreach (var element in input.ChildElements)
            {
                if (nodeCount.ContainsKey(element.LocalName))
                {
                    nodeCount[element.LocalName]++;
                }
                else
                {
                    nodeCount.Add(element.LocalName, 0);
                }
                
                output.Children.Add(ConvertToTreeNode(element, output.Path, nodeCount[element.LocalName]));
            }
            return output;
        }

        private List<string> GetProperties(DocumentTreeNode input)
        {
            if (input.Tag == null)
            {
                return new List<string>();
            }
            
            var elements = input.Tag.GetAttributes();
            var output = new List<string>(elements.Count);
            var element = input.Tag;
            foreach(var attribute in elements)
            {
                output.Add(attribute.LocalName);
            }
            if (!string.IsNullOrEmpty(element.InnerText))
            {
                output.Add("InnerText");
            }

            return output;
        }
    }

    public class DocumentTreeNode
    {
        public ObservableCollection<DocumentTreeNode> Children { get; set; }
        public string Text { get; set; }
        public string Path { get; set; }
        public OpenXmlElement? Tag { get; set; }
        

        public DocumentTreeNode(OpenXmlElement tag, int occurenceNumber, string path)
        {
            Text = occurenceNumber == 0 ? tag.LocalName : $"{tag.LocalName}[{occurenceNumber}]";
            Tag = tag;
            Path = $"{path}/{Text}";
            Children = new ObservableCollection<DocumentTreeNode>();
        }

        public DocumentTreeNode(string text)
        {
            Text = text;
            Children = new ObservableCollection<DocumentTreeNode>();
        }
    }
    
}