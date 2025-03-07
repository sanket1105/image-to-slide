import base64
import json
import os
from io import BytesIO

import requests
import streamlit as st
from dotenv import load_dotenv
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

try:
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
except ImportError:
    MSO_LINE_DASH_STYLE = None
from pptx.util import Inches, Pt

# Load environment variables (e.g., API keys)
load_dotenv()


class WhiteboardToPPT:
    def __init__(self):
        # Use environment variable or Streamlit secrets for the OpenAI API key
        self.vision_api_key = os.getenv("OPENAI_API_KEY")
        self.vision_api_url = "https://api.openai.com/v1/chat/completions"
        self.slide_width = Inches(10)
        self.slide_height = Inches(7.5)

    @st.experimental_memo(show_spinner=False)
    def analyze_image(self, image_data):
        """
        Use AI vision model to analyze the whiteboard image and extract information.
        Returns structured JSON data containing phases, sections, elements, and connections.
        This function is cached to prevent reprocessing the same image.
        """
        try:
            image_base64 = base64.b64encode(image_data).decode("utf-8")
            prompt = """
            Analyze this whiteboard image and extract:
            1. All text content with exact wording (avoid duplications)
            2. The spatial layout (what text is in boxes, what's connected to what)
            3. The hierarchical structure (main sections, subsections)
            
            Format your response as JSON with these keys:
            - phases: List of main phase headings at the top (e.g., ["Ideation", "Align", "Launch", "Manage"])
            - sections: List of vertical sections on the left side
            - elements: List of objects for each element with properties:
              {
                "id": "unique element ID (e.g., elem1)",
                "text": "Exact text content",
                "position": "top-left"|"top-center"|...,
                "type": "box"|"text"|"note",
                "coordinates": {"x": 20, "y": 30},
                "has_shape": true|false
              }
            - connections: List of objects describing logical connections between elements:
              {
                "from_id": "ID of source element",
                "to_id": "ID of target element"
              }
            
            Important:
            - Avoid duplicate texts in elements.
            - Represent duplicate texts as a single element.
            - Ensure each element has a unique ID.
            - Carefully capture all arrows/lines in the image as connections.
            """
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.vision_api_key}",
            }
            payload = {
                "model": "gpt-4o",  # Update to your available model if needed
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{image_base64}"
                                },
                            },
                        ],
                    }
                ],
                "max_tokens": 4000,
            }
            st.write("Analyzing image with AI vision model...")
            response = requests.post(self.vision_api_url, headers=headers, json=payload)
            if response.status_code != 200:
                st.error(f"API Error: {response.status_code}")
                st.error(response.text)
                return None

            result = response.json()
            if "choices" not in result or len(result["choices"]) == 0:
                st.error("Unexpected API response format")
                st.write(result)
                return None

            content = result["choices"][0]["message"]["content"]

            try:
                json_content = self._extract_json_from_text(content)
                analysis_data = json.loads(json_content)
                # Log detected connections for user feedback
                if "connections" in analysis_data:
                    st.write(
                        f"Detected {len(analysis_data['connections'])} connection(s):"
                    )
                    for conn in analysis_data["connections"]:
                        st.write(f"{conn.get('from_id')} → {conn.get('to_id')}")
                return analysis_data
            except json.JSONDecodeError as e:
                st.error(f"Error parsing JSON response: {e}")
                st.write("Raw content:")
                st.write(content)
                return self._fallback_extraction(content)
        except Exception as e:
            st.error(f"Error analyzing image: {e}")
            return None

    def _extract_json_from_text(self, text):
        """Extract JSON content from mixed text output."""
        start_idx = text.find("{")
        end_idx = text.rfind("}")
        if start_idx >= 0 and end_idx > start_idx:
            return text[start_idx : end_idx + 1]
        return "{}"

    def _fallback_extraction(self, content):
        """Fallback extraction method if JSON parsing fails."""
        st.warning("Using fallback extraction method")
        return {
            "phases": ["Ideation", "Align", "Launch", "Manage"],
            "sections": ["Strategy", "Planning", "Foundation"],
            "elements": [
                {
                    "id": "elem1",
                    "text": "Content Ideas",
                    "position": "middle-left",
                    "type": "box",
                    "coordinates": {"x": 20, "y": 30},
                    "has_shape": True,
                },
            ],
            "connections": [{"from_id": "elem1", "to_id": "elem1"}],
        }

    def adjust_positions_for_overlap(self, elements_dict):
        """Adjust positions to prevent overlaps between elements."""
        elems = list(elements_dict.items())
        for i in range(len(elems)):
            for j in range(i + 1, len(elems)):
                id1, elem1 = elems[i]
                id2, elem2 = elems[j]
                right1 = elem1["left"] + elem1["width"]
                bottom1 = elem1["top"] + elem1["height"]
                right2 = elem2["left"] + elem2["width"]
                bottom2 = elem2["top"] + elem2["height"]
                if (
                    elem1["left"] < right2
                    and right1 > elem2["left"]
                    and elem1["top"] < bottom2
                    and bottom1 > elem2["top"]
                ):
                    # Simple adjustment: move the second element down by 0.2 inches
                    elements_dict[id2]["top"] += Inches(0.2)
        return elements_dict

    def create_ppt_from_analysis(self, analysis_data):
        """
        Create a PowerPoint slide based on the AI-analyzed content.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 1. Add phase headers at the top
        if "phases" in analysis_data:
            self._add_phases(slide, analysis_data["phases"])

        # 2. Add vertical section labels on the left
        if "sections" in analysis_data:
            self._add_sections(slide, analysis_data["sections"])

        # 3. Build a mapping of positions; use provided coordinates if available.
        position_map = {
            "top-left": {"x": 1.2, "y": 1.0},
            "top-center": {"x": 4.0, "y": 1.0},
            "top-right": {"x": 7.0, "y": 1.0},
            "middle-left": {"x": 1.2, "y": 2.8},
            "middle-center": {"x": 4.0, "y": 2.8},
            "middle-right": {"x": 7.0, "y": 2.8},
            "bottom-left": {"x": 1.2, "y": 4.6},
            "bottom-center": {"x": 4.0, "y": 4.6},
            "bottom-right": {"x": 7.0, "y": 4.6},
        }

        # Store element shapes by their unique ID
        element_shapes = {}

        # Add dynamic positions based on provided coordinates
        if "elements" in analysis_data:
            for element in analysis_data["elements"]:
                if "coordinates" in element and "id" in element:
                    elem_id = element["id"]
                    # Convert percentages to inches (scaled within available area)
                    x_in = max(1.2, min(8.5, (element["coordinates"]["x"] / 100) * 9))
                    y_in = max(1.0, min(6.5, (element["coordinates"]["y"] / 100) * 6.5))
                    position_map[elem_id] = {"x": x_in, "y": y_in}
                    element["position_key"] = elem_id

        # 4. Add elements (boxes) and store their shape info
        if "elements" in analysis_data:
            for element in analysis_data["elements"]:
                elem_id = element.get("id", f"elem_{id(element)}")
                position = element.get(
                    "position_key", element.get("position", "middle-center")
                )
                pos_config = position_map.get(position, {"x": 4.0, "y": 2.8})
                left = Inches(pos_config["x"])
                top = Inches(pos_config["y"])
                text_str = element["text"]
                text_len = len(text_str)
                estimated_lines = max(1, text_len // 25)
                width = min(Inches(3.5), max(Inches(1.5), Inches(text_len / 12)))
                height = min(
                    Inches(1.2), max(Inches(0.5), Inches(estimated_lines * 0.3))
                )

                shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, left, top, width, height
                )
                shape.line.color.rgb = RGBColor(0, 0, 0)
                shape.fill.background()
                tf = shape.text_frame
                tf.text = text_str
                tf.word_wrap = True

                # Adjust font size based on text length
                if text_len > 70:
                    font_size = Pt(8)
                elif text_len > 50:
                    font_size = Pt(9)
                elif text_len > 30:
                    font_size = Pt(10)
                else:
                    font_size = Pt(12)
                p = tf.paragraphs[0]
                p.font.size = font_size
                p.font.name = "Arial"
                p.alignment = PP_ALIGN.CENTER

                element_shapes[elem_id] = {
                    "shape": shape,
                    "left": left,
                    "top": top,
                    "width": width,
                    "height": height,
                }

        # Adjust positions to prevent overlaps
        element_shapes = self.adjust_positions_for_overlap(element_shapes)
        for elem_id, info in element_shapes.items():
            info["shape"].left = info["left"]
            info["shape"].top = info["top"]

        # 5. Draw connection arrows based on detected connections
        if "connections" in analysis_data:
            for connection in analysis_data["connections"]:
                from_id = connection.get("from_id")
                to_id = connection.get("to_id")
                if from_id not in element_shapes or to_id not in element_shapes:
                    continue

                from_sh = element_shapes[from_id]
                to_sh = element_shapes[to_id]
                # Calculate centers for connection endpoints
                from_center_x = from_sh["left"] + (from_sh["width"] / 2)
                from_center_y = from_sh["top"] + (from_sh["height"] / 2)
                to_center_x = to_sh["left"] + (to_sh["width"] / 2)
                to_center_y = to_sh["top"] + (to_sh["height"] / 2)

                # Determine best connection direction (horizontal or vertical)
                if abs(from_center_x - to_center_x) > abs(from_center_y - to_center_y):
                    if from_center_x < to_center_x:
                        start_x = from_sh["left"] + from_sh["width"]
                        start_y = from_center_y
                        end_x = to_sh["left"]
                        end_y = to_center_y
                    else:
                        start_x = from_sh["left"]
                        start_y = from_center_y
                        end_x = to_sh["left"] + to_sh["width"]
                        end_y = to_center_y
                else:
                    if from_center_y < to_center_y:
                        start_x = from_center_x
                        start_y = from_sh["top"] + from_sh["height"]
                        end_x = to_center_x
                        end_y = to_sh["top"]
                    else:
                        start_x = from_center_x
                        start_y = from_sh["top"]
                        end_x = to_center_x
                        end_y = to_sh["top"] + to_sh["height"]

                connector = slide.shapes.add_connector(
                    MSO_CONNECTOR.STRAIGHT, start_x, start_y, end_x, end_y
                )
                connector.line.color.rgb = RGBColor(0, 0, 0)
                connector.line.width = Pt(1.5)
                if MSO_LINE_DASH_STYLE:
                    connector.line.dash_style = MSO_LINE_DASH_STYLE.SOLID
                # Set basic arrowheads using integer values (0 = none, 3 = triangle)
                connector.line.begin_arrowhead = 0
                connector.line.end_arrowhead = 3

        pptx_io = BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)
        return pptx_io

    def process_image_to_ppt(self, image_data):
        """
        Main method: Analyze the image and create a PowerPoint slide.
        """
        analysis_data = self.analyze_image(image_data)
        if not analysis_data:
            st.error("Failed to analyze image")
            return None
        st.write("Creating PowerPoint slide...")
        pptx_data = self.create_ppt_from_analysis(analysis_data)
        return pptx_data


def main():
    st.set_page_config(page_title="Whiteboard to PowerPoint Converter", page_icon="📊")
    st.title("Whiteboard to PowerPoint Converter")
    st.write(
        "Upload a whiteboard image and get a structured PowerPoint slide with elements and logical connections."
    )

    # Ensure API key is provided
    api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        st.warning("⚠️ OpenAI API key not found. Enter it below:")
        api_key = st.text_input("Enter your OpenAI API key:", type="password")
        if api_key:
            os.environ["OPENAI_API_KEY"] = api_key

    uploaded_file = st.file_uploader(
        "Upload whiteboard image", type=["jpg", "jpeg", "png"]
    )
    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="Uploaded Whiteboard Image", use_container_width=True)

        if st.button("Convert to PowerPoint"):
            with st.spinner("Processing..."):
                uploaded_file.seek(0)
                image_data = uploaded_file.read()

                converter = WhiteboardToPPT()
                pptx_data = converter.process_image_to_ppt(image_data)

                if pptx_data:
                    st.success("PowerPoint slide created successfully!")
                    st.download_button(
                        label="Download PowerPoint",
                        data=pptx_data,
                        file_name="whiteboard_to_ppt.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
                else:
                    st.error(
                        "Failed to create PowerPoint. Please check the logs above for details."
                    )


if __name__ == "__main__":
    main()
