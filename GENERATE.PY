
import os
import json
import argparse
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--topic", required=True)
    parser.add_argument("--slides", type=int, default=10)
    parser.add_argument("--language", default="Uzbek")
    args = parser.parse_args()

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY topilmadi!")

    client = genai.Client(api_key=api_key)

    prompt = f"""
    Create a professional presentation in {args.language}.
    Topic: {args.topic}
    Number of slides: {args.slides}

    Return ONLY valid JSON in this format:
    {{
      "title": "Presentation title",
      "slides": [
        {{
          "title": "Slide title",
          "bullets": ["Point 1", "Point 2", "Point 3"]
        }}
      ]
    }}
    Make content informative, accurate and detailed.
    """

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config={"response_mime_type": "application/json"}
    )

    data = json.loads(response.text)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for index, item in enumerate(data["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = RGBColor(20, 35, 65)

        title_box = slide.shapes.add_textbox(
            Inches(0.7), Inches(0.7),
            Inches(12), Inches(1.2)
        )
        title = title_box.text_frame.paragraphs[0]
        title.text = item["title"]
        title.font.size = Pt(30)
        title.font.bold = True
        title.font.color.rgb = RGBColor(255, 210, 90)

        body = slide.shapes.add_textbox(
            Inches(1), Inches(2.2),
            Inches(11.5), Inches(4.5)
        )
        tf = body.text_frame
        tf.word_wrap = True

        for bullet in item["bullets"]:
            p = tf.add_paragraph()
            p.text = "• " + bullet
            p.font.size = Pt(21)
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.space_after = Pt(18)

    prs.save("presentation.pptx")
    print("Presentation created successfully!")

if __name__ == "__main__":
    main()
