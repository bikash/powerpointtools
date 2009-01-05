using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLazySlides
{
    class PresentationTags
    {
        public AddInTagInt LazySlideIdCounter;

        public PresentationTags(Presentation presentation)
        {
            Tags tags = presentation.Tags;

            LazySlideIdCounter = new AddInTagInt(tags, "LazySlideIdCounter");
        }

        public void Clear()
        {
            LazySlideIdCounter.Clear();
        }
    }

    class SlideTags
    {
        public AddInTagBool IsLazySlide;
        public AddInTagBool HasLazySlide;
        public AddInTagInt LazySlideId;
        public AddInTagInt ParentLazySlideId;
        public AddInTagInt ChildLazySlideId;

        public SlideTags(Slide slide)
        {
            Tags tags = slide.Tags;

            IsLazySlide = new AddInTagBool(tags, "IsLazySlide");
            HasLazySlide = new AddInTagBool(tags, "HasLazySlide");

            LazySlideId = new AddInTagInt(tags, "LazySlideId");

            ParentLazySlideId = new AddInTagInt(tags, "ParentLazySlideId");
            ChildLazySlideId = new AddInTagInt(tags, "ChildLazySlideId");
        }

        public void Clear()
        {
            IsLazySlide.Clear();
            HasLazySlide.Clear();

            LazySlideId.Clear();

            ParentLazySlideId.Clear();
            ChildLazySlideId.Clear();
        }
    }

    class ShapeTags
    {
        public AddInTagBool ExcludeParentText;
        public AddInTagInt ParentTextLength;

        public AddInTagInt LazySlideId;

        public AddInTagInt ParentLazySlideId;
        public AddInTagInt ChildLazySlideId;

        public ShapeTags(Shape shape)
        {
            Tags tags = shape.Tags;

            ExcludeParentText = new AddInTagBool(tags, "ExcludeParentText");
            ParentTextLength = new AddInTagInt(tags, "ParentTextLength");

            LazySlideId = new AddInTagInt(tags, "LazySlideId");

            ParentLazySlideId = new AddInTagInt(tags, "ParentLazySlideId");
            ChildLazySlideId = new AddInTagInt(tags, "ChildLazySlideId");
        }

        public void Clear()
        {
            ExcludeParentText.Clear();
            ParentTextLength.Clear();

            LazySlideId.Clear();

            ParentLazySlideId.Clear();
            ChildLazySlideId.Clear();
        }
    }

    static class Extensions
    {
        public static PresentationTags LazySlideTags(this Presentation presentation)
        {
            return new PresentationTags(presentation);
        }

        public static SlideTags LazySlideTags(this Slide slide)
        {
            return new SlideTags(slide);
        }

        public static ShapeTags LazySlideTags(this Shape shape)
        {
            return new ShapeTags(shape);
        }
    }
}
