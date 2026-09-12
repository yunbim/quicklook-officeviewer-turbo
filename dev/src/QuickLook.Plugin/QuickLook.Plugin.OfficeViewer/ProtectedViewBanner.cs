// Copyright © 2017-2026 QL-Win Contributors
//
// This file is part of QuickLook program.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace QuickLook.Plugin.OfficeViewer;

/// <summary>
/// The slim notice strip placed above a preview that was rendered from a temporary copy of a
/// Protected View document (see <see cref="ProtectedViewPreview"/>).
///
/// <para>It does two things the old confirmation dialog could not: it states up front that what
/// is on screen is a copy and that the original file was left alone, and it offers an explicit,
/// user-triggered way to remove the Protected View mark from the original — so the decision to
/// modify a file is always made knowingly, and never as a side effect of looking at it.</para>
///
/// <para>The strip is deliberately built from plain shapes with hard-coded colours instead of
/// theme resources. A <c>WindowsFormsHost</c> always paints over its WPF siblings, so the notice
/// has to live in its own layout row rather than float above the preview — and because it sits
/// in the layout, its appearance must not depend on whether QuickLook is running light or dark.</para>
/// </summary>
internal sealed class ProtectedViewBanner : Grid
{
    private static readonly Brush StripBackground = Freeze(Color.FromRgb(0xFF, 0xF4, 0xCE));
    private static readonly Brush StripBorder = Freeze(Color.FromRgb(0xE3, 0xC0, 0x6A));
    private static readonly Brush PrimaryText = Freeze(Color.FromRgb(0x5C, 0x44, 0x00));
    private static readonly Brush ButtonFace = Freeze(Colors.White);
    private static readonly Brush ButtonFaceHover = Freeze(Color.FromRgb(0xFF, 0xE6, 0x9C));
    private static readonly Brush ButtonEdge = Freeze(Color.FromRgb(0xC9, 0xA2, 0x27));

    private static readonly Brush SuccessBackground = Freeze(Color.FromRgb(0xE7, 0xF6, 0xE9));
    private static readonly Brush SuccessBorder = Freeze(Color.FromRgb(0xA8, 0xD5, 0xAE));
    private static readonly Brush SuccessText = Freeze(Color.FromRgb(0x1E, 0x5B, 0x2B));

    private static readonly Brush FailureBackground = Freeze(Color.FromRgb(0xFD, 0xE7, 0xE9));
    private static readonly Brush FailureBorder = Freeze(Color.FromRgb(0xE7, 0xA9, 0xAF));
    private static readonly Brush FailureText = Freeze(Color.FromRgb(0x8E, 0x1F, 0x28));

    private readonly string _sourcePath;
    private readonly Border _strip;
    private readonly TextBlock _message;
    private readonly ContentControl _actionHost;

    public ProtectedViewBanner(string sourcePath, UIElement content)
    {
        _sourcePath = sourcePath;

        RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        _message = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = PrimaryText,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Text = UiText.BannerMessage,
        };

        _actionHost = new ContentControl { VerticalAlignment = VerticalAlignment.Center };
        _actionHost.Content = BuildActionButton(UiText.BannerActionLabel, UnblockSource);

        var dock = new DockPanel { LastChildFill = true };
        DockPanel.SetDock(_actionHost, Dock.Right);
        dock.Children.Add(_actionHost);
        dock.Children.Add(_message);

        _strip = new Border
        {
            Background = StripBackground,
            BorderBrush = StripBorder,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(12, 7, 12, 7),
            Child = dock,
        };

        Grid.SetRow(_strip, 0);
        Children.Add(_strip);

        Grid.SetRow(content, 1);
        Children.Add(content);
    }

    /// <summary>
    /// Removes the Protected View mark from the original file. Runs on the UI thread and only
    /// in response to a click, so the user is always the one deciding to change the file.
    /// </summary>
    private void UnblockSource()
    {
        var unblocked = ProtectedViewPreview.UnblockSource(_sourcePath);

        if (unblocked)
        {
            _strip.Background = SuccessBackground;
            _strip.BorderBrush = SuccessBorder;
            _message.Foreground = SuccessText;
            _message.Text = UiText.BannerDoneMessage;
            _actionHost.Content = BuildDoneTag();
        }
        else
        {
            _strip.Background = FailureBackground;
            _strip.BorderBrush = FailureBorder;
            _message.Foreground = FailureText;
            _message.Text = UiText.BannerFailureMessage;
            _actionHost.Content = BuildActionButton(UiText.BannerRetryLabel, UnblockSource);
        }
    }

    private static Border BuildActionButton(string caption, Action onClick)
    {
        var label = new TextBlock
        {
            Text = caption,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = PrimaryText,
            VerticalAlignment = VerticalAlignment.Center,
        };

        var button = new Border
        {
            Child = label,
            Background = ButtonFace,
            BorderBrush = ButtonEdge,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(12, 4, 12, 4),
            Margin = new Thickness(14, 0, 0, 0),
            Cursor = Cursors.Hand,
            ToolTip = UiText.BannerActionTip,
        };

        button.MouseEnter += (_, __) => button.Background = ButtonFaceHover;
        button.MouseLeave += (_, __) => button.Background = ButtonFace;
        button.MouseLeftButtonUp += (_, __) => onClick();

        return button;
    }

    private static TextBlock BuildDoneTag()
    {
        return new TextBlock
        {
            Text = UiText.BannerDoneTag,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = SuccessText,
            Margin = new Thickness(14, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
        };
    }

    private static Brush Freeze(Color color)
    {
        var brush = new SolidColorBrush(color);
        brush.Freeze();
        return brush;
    }
}
