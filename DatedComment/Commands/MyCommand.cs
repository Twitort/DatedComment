using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System.Diagnostics.Eventing.Reader;
using System.Linq;

namespace DatedComment
{
    [Command(PackageIds.MyCommand)]
    internal sealed class MyCommand : BaseCommand<MyCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            // Get the current doc being edited:
            var docView = await VS.Documents.GetActiveDocumentViewAsync();

            // Get the current first text selection:
            var curSelection = docView?.TextView.Selection.SelectedSpans.FirstOrDefault();

            if (curSelection.HasValue) 
            {
                // Get the white space to the left of the caret so that
                // it can be inserted on subsequent lines:
                string whiteSpace = GetLeftWhiteSpace(docView);

                var commentBlock = @"// ";

                // Drop the comment into the editor, replacing the current selection if any:
                var theBufr = docView.TextBuffer;
                docView.TextBuffer.Replace(curSelection.Value, commentBlock);

                curSelection = docView?.TextView.Selection.SelectedSpans.FirstOrDefault();

                if (curSelection.HasValue)
                {
                    // Build the dated comment block:
                    var restOfBlock = Environment.NewLine + whiteSpace + @"// " + Environment.UserName + " - " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine + whiteSpace;

                    // Drop the comment into the editor, replacing the current selection if any:
                    docView.TextBuffer.Replace(curSelection.Value, restOfBlock);
                    var caretPosition = curSelection.Value.Start;

                    var textViewCaret = docView.TextView.Caret;
                    textViewCaret.MoveTo(new SnapshotPoint(docView.TextBuffer.CurrentSnapshot, caretPosition));
                }
            }
        }

        private string GetLeftWhiteSpace(DocumentView docView)
        {
            // Get the text buffer's text snapshot:
            ITextSnapshot snapshot = docView.TextBuffer.CurrentSnapshot;

            // Get the current caret position:
            ITextView textView = docView.TextView;
            SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;

            // Note the current spot and walk to the left until a newline is hit or
            // beginning of document:
            int position = caretPosition.Position;
            int whitespaceStart = position;
            while (whitespaceStart > 0 && IsTabOrSpace(snapshot.GetText(--whitespaceStart, 1)[0])) ;

            // Pull the text between the white space start and the current position:
            int whitespaceLength = position - ++whitespaceStart;
            return (whitespaceLength > 0) ? snapshot.GetText(whitespaceStart, whitespaceLength) : "";
        }

        private bool IsTabOrSpace(char c)
        {
            return (c == ' ' || c == '\t');
        }
    }
}
