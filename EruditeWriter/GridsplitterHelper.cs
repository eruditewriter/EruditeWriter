using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace EruditeWriter
{
    public class FixedGridSplitter : GridSplitter
    {
        private Grid grid;
        private ColumnDefinition definition1;
        private double savedMaxLength;

        #region static

        static FixedGridSplitter()
        {
            new GridSplitter();
            EventManager.RegisterClassHandler(typeof(FixedGridSplitter), Thumb.DragCompletedEvent, new DragCompletedEventHandler(FixedGridSplitter.OnDragCompleted));
            EventManager.RegisterClassHandler(typeof(FixedGridSplitter), Thumb.DragStartedEvent, new DragStartedEventHandler(FixedGridSplitter.OnDragStarted));
        }

        private static void OnDragStarted(object sender, DragStartedEventArgs e)
        {
            FixedGridSplitter splitter = (FixedGridSplitter)sender;
            splitter.OnDragStarted(e);
        }

        private static void OnDragCompleted(object sender, DragCompletedEventArgs e)
        {
            FixedGridSplitter splitter = (FixedGridSplitter)sender;
            splitter.OnDragCompleted(e);
        }

        #endregion

        private void OnDragStarted(DragStartedEventArgs sender)
        {
            grid = Parent as Grid;
            if (grid == null)
                return;
            int splitterIndex = (int)GetValue(Grid.ColumnProperty);
            definition1 = grid.ColumnDefinitions[splitterIndex - 1];
            ColumnDefinition definition2 = grid.ColumnDefinitions[splitterIndex + 1];
            savedMaxLength = definition1.MaxWidth;
            double maxWidth = definition1.ActualWidth + definition2.ActualWidth - definition2.MinWidth;
            definition1.MaxWidth = maxWidth;
        }

        private void OnDragCompleted(DragCompletedEventArgs sender)
        {
            definition1.MaxWidth = savedMaxLength;
            grid = null;
            definition1 = null;
        }
    }
}