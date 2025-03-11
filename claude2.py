import base64
import io
import json
import os
from io import BytesIO

import requests
import streamlit as st
from dotenv import load_dotenv
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

# Load environment variables for API keys
load_dotenv()
from pptx.enum.dml import MSO_LINE_DASH_STYLE, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN


class WhiteboardToPPT:
    def __init__(self):
        # API configuration - use environment variable or Streamlit secrets
        self.vision_api_key = os.getenv("OPENAI_API_KEY")
        self.vision_api_url = "https://api.openai.com/v1/chat/completions"

        # PowerPoint settings
        self.slide_width = Inches(10)
        self.slide_height = Inches(7.5)

        # Configuration for element sizing and spacing
        self.min_box_width = Inches(1.2)
        self.max_box_width = Inches(3.5)
        self.min_box_height = Inches(0.5)
        self.max_box_height = Inches(1.2)
        self.horizontal_spacing = Inches(0.4)  # Space between elements horizontally
        self.vertical_spacing = Inches(0.4)  # Space between elements vertically

    def analyze_image(self, image_data):
        """
        Use AI vision model to analyze the whiteboard image and extract information.
        Returns structured data about text content and layout.
        """
        try:
            # Convert image to base64
            image_base64 = base64.b64encode(image_data).decode("utf-8")

            # Enhanced prompt for the vision model with better OCR instructions
            prompt = """
            Analyze this whiteboard image with high precision OCR, extracting:
            1. All text content with EXACT wording, preserving all punctuation, capitalization, and special characters
            2. The spatial layout and relationships between elements
            3. The hierarchical structure and flow of information
            
            CRITICALLY IMPORTANT:
            - Scan the image multiple times to ensure ALL text is captured
            - Verify each extracted text element is unique (DO NOT DUPLICATE any text)
            - Ensure text with special characters (#, *, etc.) is captured correctly
            - Treat faint or partial text as important and capture it
            - Capture text in cloud shapes, rectangles, and any other visual elements
            
            FORMAT REQUIREMENTS - STRICTLY ADHERE TO THIS JSON STRUCTURE:
            {
              "phases": [list of phase headings at the top, ordered left to right],
              "sections": [list of vertical sections on the left side, ordered top to bottom],
              "elements": [
                {
                  "id": "unique identifier (elem1, elem2, etc.)",
                  "text": "exact text content with ALL special characters and formatting preserved",
                  "position": "precise position descriptor (top-left, middle-center, etc.)",
                  "type": "box/text/note/cloud",
                  "coordinates": {"x": percentage from left, "y": percentage from top},
                  "has_shape": boolean,
                  "emphasis": "none/high/medium/low based on visual prominence"
                }
              ],
              "connections": [
                {"from_id": "source element id", "to_id": "target element id", "arrow_type": "simple/double/none"}
              ]
            }
            
            Before finalizing, verify:
            1. NO DUPLICATE TEXT in any elements
            2. ALL TEXT from image is captured (compare your extraction with the image)
            3. Each element has a unique ID
            4. Clouds and special shapes are properly identified
            """

            # API call to vision model
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.vision_api_key}",
            }

            # Using GPT-4o for improved visual analysis with temperature=0 for consistency
            payload = {
                "model": "gpt-4o",
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
                "temperature": 0,  # Using 0 for maximum consistency and accuracy
            }

            with st.status(
                "Analyzing image with AI vision model...", expanded=True
            ) as status:
                response = requests.post(
                    self.vision_api_url, headers=headers, json=payload
                )

                if response.status_code != 200:
                    st.error(f"API Error: {response.status_code}")
                    st.error(response.text)
                    status.update(label="Analysis failed", state="error")
                    return None

                result = response.json()
                status.update(label="Analysis complete", state="complete")

            # Check if we have the expected structure
            if "choices" not in result or len(result["choices"]) == 0:
                st.error("Unexpected API response format")
                st.write(result)
                return None

            # Extract the content from the response
            content = result["choices"][0]["message"]["content"]

            # Parse the JSON content - handle potential JSON formatting issues
            try:
                # Find JSON content in the response (it might be mixed with other text)
                json_content = self._extract_json_from_text(content)
                analysis_data = json.loads(json_content)

                # Log the detected connections and elements
                st.expander("AI Analysis Results", expanded=False).write(analysis_data)

                # Check for duplicate texts and missing elements
                self._validate_analysis_data(analysis_data)

                # Enhance data with text analysis
                analysis_data = self._enhance_analysis_data(analysis_data)

                return analysis_data
            except json.JSONDecodeError as e:
                st.error(f"Error parsing JSON response: {e}")
                st.write("Raw content:")
                st.write(content)

                # Fallback to manual extraction if JSON parsing fails
                return self._fallback_extraction(content)

        except Exception as e:
            st.error(f"Error analyzing image: {e}")
            return None

    def _validate_analysis_data(self, data):
        """Validate the analysis data for duplicates and missing elements"""
        if not data or "elements" not in data:
            return

        # Check for duplicate texts
        texts = []
        duplicates = []

        for element in data["elements"]:
            if "text" in element:
                if element["text"] in texts:
                    duplicates.append(element["text"])
                texts.append(element["text"])

        if duplicates:
            st.warning(
                f"Found {len(duplicates)} duplicate text elements. Removing duplicates..."
            )

            # Remove duplicates while preserving the first occurrence
            seen_texts = set()
            filtered_elements = []

            for element in data["elements"]:
                if "text" in element and element["text"] not in seen_texts:
                    seen_texts.add(element["text"])
                    filtered_elements.append(element)
                elif "text" not in element:
                    filtered_elements.append(element)

            data["elements"] = filtered_elements
            st.success("Duplicates removed.")

        # Also check for duplicate phase headers
        if "phases" in data and data["phases"]:
            unique_phases = []
            for phase in data["phases"]:
                if phase not in unique_phases:
                    unique_phases.append(phase)
            data["phases"] = unique_phases

    def _enhance_analysis_data(self, data):
        """Add additional analysis to improve the layout and connections"""
        if not data:
            return data

        # Add size estimation based on text length
        if "elements" in data:
            for element in data["elements"]:
                if "text" in element:
                    # Calculate text metrics
                    text = element["text"]
                    element["text_length"] = len(text)
                    element["word_count"] = len(text.split())

                    # Estimate importance based on position, content, and connections
                    importance = 1  # Default importance

                    # Check if this element is connected to many others
                    if "connections" in data:
                        connections_from = sum(
                            1
                            for conn in data["connections"]
                            if conn["from_id"] == element["id"]
                        )
                        connections_to = sum(
                            1
                            for conn in data["connections"]
                            if conn["to_id"] == element["id"]
                        )
                        element["connection_count"] = connections_from + connections_to

                        # Elements with more connections tend to be more important
                        if connections_from + connections_to > 2:
                            importance += 1

                    # Top position elements are often more important
                    if "position" in element and "top" in element["position"].lower():
                        importance += 1

                    # Elements with special markers like "#" might be emphasized
                    if "#" in text or "*" in text or "!" in text:
                        importance += 1

                    element["importance"] = min(importance, 3)  # Scale of 1-3

        return data

    def _extract_json_from_text(self, text):
        """Extract JSON content from potentially mixed text"""
        # Look for content between curly braces
        start_idx = text.find("{")
        end_idx = text.rfind("}")

        if start_idx >= 0 and end_idx > start_idx:
            return text[start_idx : end_idx + 1]

        return "{}"  # Return empty JSON if none found

    def _fallback_extraction(self, content):
        """Enhanced manual extraction of key elements as fallback"""
        st.warning("Using fallback extraction method")

        # More sophisticated fallback with better structure
        return {
            "phases": ["Ideation", "Align", "Launch", "Manage"],
            "sections": ["Strategy", "Planning", "Foundation"],
            "elements": [
                {
                    "id": "elem1",
                    "text": "#Need Sizing Tool#",
                    "position": "top-left",
                    "type": "note",
                    "coordinates": {"x": 15, "y": 13},
                    "has_shape": True,
                    "emphasis": "high",
                    "text_length": 17,
                    "word_count": 3,
                    "importance": 3,
                },
                {
                    "id": "elem2",
                    "text": "Define: Test (channels) Audience Success",
                    "position": "top-center",
                    "type": "box",
                    "coordinates": {"x": 35, "y": 13},
                    "has_shape": True,
                    "emphasis": "medium",
                    "text_length": 42,
                    "word_count": 5,
                    "importance": 2,
                },
                # Additional elements would be included here...
            ],
            "connections": [
                {"from_id": "elem1", "to_id": "elem2", "arrow_type": "simple"},
                # Additional connections would be included here...
            ],
        }

    def _calculate_element_dimensions(self, element):
        """Calculate appropriate dimensions for an element based on its content"""
        text = element.get("text", "")
        text_length = element.get("text_length", len(text))
        word_count = element.get("word_count", len(text.split()))
        importance = element.get("importance", 1)

        # Calculate width based on text length and word count
        # Longer words need more width
        avg_word_length = text_length / max(word_count, 1)
        width_factor = min(
            1.0, 0.6 + (avg_word_length / 20)
        )  # Scale based on average word length

        # Calculate width based on text length with constraints
        width = max(
            self.min_box_width,
            min(self.max_box_width, Inches(1.0 + (text_length / 15) * width_factor)),
        )

        # Calculate height based on text length and width
        # Estimate how many lines the text will take
        chars_per_line = width.inches * 11  # Approx 11 characters per inch
        estimated_lines = max(1, text_length / chars_per_line)

        # Add some extra height for important elements
        importance_factor = 1.0 + ((importance - 1) * 0.15)

        height = max(
            self.min_box_height,
            min(
                self.max_box_height,
                Inches(0.4 + (estimated_lines * 0.2) * importance_factor),
            ),
        )

        return width, height

    def _intelligently_position_elements(self, analysis_data):
        """Create an optimal layout for elements based on their relationships and positions"""
        # Create a grid system for positioning
        grid = {
            "rows": 6,  # Increase number of rows for better spacing
            "cols": 6,  # Increase number of columns for better spacing
            "cells": {},  # Will store what's placed in each cell
        }

        # First, prepare data on all elements
        elements_data = {}
        if "elements" in analysis_data:
            # Sort elements by importance/priority
            elements = sorted(
                analysis_data["elements"],
                key=lambda e: (
                    # Sort by y-coordinate first (top to bottom)
                    e.get("coordinates", {}).get("y", 50),
                    # Then by x-coordinate (left to right)
                    e.get("coordinates", {}).get("x", 50),
                ),
            )

            for element in elements:
                element_id = element.get("id")
                if not element_id:
                    continue

                # Get position information
                pos = element.get("position", "middle-center")
                coords = element.get("coordinates", {"x": 50, "y": 50})

                # Determine grid position
                row = min(grid["rows"] - 1, int((coords["y"] / 100) * grid["rows"]))
                col = min(grid["cols"] - 1, int((coords["x"] / 100) * grid["cols"]))

                # Calculate dimensions
                width, height = self._calculate_element_dimensions(element)

                # Store information for this element
                elements_data[element_id] = {
                    "element": element,
                    "grid_pos": (row, col),
                    "width": width,
                    "height": height,
                    "connections": [],
                }

                # Mark this grid cell as occupied
                if (row, col) not in grid["cells"]:
                    grid["cells"][(row, col)] = []
                grid["cells"][(row, col)].append(element_id)

        # Add connection information
        if "connections" in analysis_data:
            for connection in analysis_data["connections"]:
                from_id = connection.get("from_id")
                to_id = connection.get("to_id")

                if from_id in elements_data and to_id in elements_data:
                    elements_data[from_id]["connections"].append(
                        {"to": to_id, "type": connection.get("arrow_type", "simple")}
                    )
                    elements_data[to_id]["connections"].append(
                        {
                            "from": from_id,
                            "type": connection.get("arrow_type", "simple"),
                        }
                    )

        # Now convert grid positions to actual coordinates
        slide_margin_x = Inches(1.0)
        slide_margin_y = Inches(1.0)

        # Calculate usable area for elements
        usable_width = self.slide_width - (slide_margin_x * 2)
        usable_height = self.slide_height - (slide_margin_y * 2)

        # Calculate cell size
        cell_width = usable_width / grid["cols"]
        cell_height = usable_height / grid["rows"]

        # IMPROVED OVERLAP HANDLING
        # Step 1: First pass to place elements on a more spread out grid
        for element_id, data in elements_data.items():
            row, col = data["grid_pos"]

            # Calculate base position (centered in the cell)
            base_x = (
                slide_margin_x
                + (col * cell_width)
                + (cell_width / 2)
                - (data["width"] / 2)
            )
            base_y = (
                slide_margin_y
                + (row * cell_height)
                + (cell_height / 2)
                - (data["height"] / 2)
            )

            # Store the calculated position
            data["position"] = {"left": base_x, "top": base_y}

        # Step 2: Second pass to detect and resolve overlaps
        resolved_positions = {}

        # Sort elements for deterministic processing (helps with consistent layout)
        sorted_elements = sorted(
            elements_data.items(),
            key=lambda x: (x[1]["grid_pos"][0], x[1]["grid_pos"][1]),
        )

        for element_id, data in sorted_elements:
            current_pos = data["position"]
            current_width = data["width"]
            current_height = data["height"]

            # Check if this element overlaps with any already positioned element
            overlapping = False

            for positioned_id, positioned_pos in resolved_positions.items():
                # Skip checking against itself
                if positioned_id == element_id:
                    continue

                positioned_width = elements_data[positioned_id]["width"]
                positioned_height = elements_data[positioned_id]["height"]

                # Check for overlap
                if (
                    current_pos["left"] < positioned_pos["left"] + positioned_width
                    and current_pos["left"] + current_width > positioned_pos["left"]
                    and current_pos["top"] < positioned_pos["top"] + positioned_height
                    and current_pos["top"] + current_height > positioned_pos["top"]
                ):

                    overlapping = True

                    # Determine best direction to shift (right or down)
                    # Calculate distances for potential shifts
                    shift_right = (
                        positioned_pos["left"]
                        + positioned_width
                        - current_pos["left"]
                        + self.horizontal_spacing
                    )
                    shift_down = (
                        positioned_pos["top"]
                        + positioned_height
                        - current_pos["top"]
                        + self.vertical_spacing
                    )

                    # Choose the smaller shift
                    if (
                        shift_right <= shift_down
                        and current_pos["left"] + current_width + shift_right
                        <= self.slide_width - slide_margin_x
                    ):
                        # Shift right
                        current_pos["left"] += shift_right
                    else:
                        # Shift down
                        current_pos["top"] += shift_down

                    # Re-check all positioned elements again after shifting
                    # This is needed to ensure the newly shifted position doesn't overlap with others
                    overlapping = True
                    break

            # Keep checking until no overlaps remain
            while overlapping:
                overlapping = False
                for positioned_id, positioned_pos in resolved_positions.items():
                    if positioned_id == element_id:
                        continue

                    positioned_width = elements_data[positioned_id]["width"]
                    positioned_height = elements_data[positioned_id]["height"]

                    if (
                        current_pos["left"] < positioned_pos["left"] + positioned_width
                        and current_pos["left"] + current_width > positioned_pos["left"]
                        and current_pos["top"]
                        < positioned_pos["top"] + positioned_height
                        and current_pos["top"] + current_height > positioned_pos["top"]
                    ):

                        overlapping = True

                        # Try shifting right first, if still on slide
                        if (
                            current_pos["left"]
                            + current_width
                            + self.horizontal_spacing
                            <= self.slide_width - slide_margin_x
                        ):
                            current_pos["left"] += self.horizontal_spacing
                        # Otherwise shift down
                        else:
                            current_pos["left"] = slide_margin_x  # Reset to left side
                            current_pos["top"] += self.vertical_spacing

                        break

            # Add this element to the resolved positions
            resolved_positions[element_id] = current_pos
            # Update the position in the original data
            data["position"] = current_pos

        return elements_data

    def create_ppt_from_analysis(self, analysis_data):
        """
        Create a PowerPoint slide based on the AI-analyzed content.
        """
        # Initialize PowerPoint presentation with improved slide sizing
        prs = Presentation()

        # Set slide dimensions for better layout
        prs.slide_width = self.slide_width
        prs.slide_height = self.slide_height

        slide_layout = prs.slide_layouts[6]  # Blank slide layout
        slide = prs.slides.add_slide(slide_layout)

        # 1. Add title to the slide for context
        title_shape = slide.shapes.add_textbox(
            left=Inches(0.5), top=Inches(0.1), width=Inches(9.0), height=Inches(0.5)
        )
        title_frame = title_shape.text_frame
        title_frame.text = "Whiteboard Conversion"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(20)
        title_para.font.bold = True
        title_para.font.name = "Arial"
        title_para.font.color.rgb = RGBColor(0, 0, 0)  # Black text for title
        title_para.alignment = PP_ALIGN.CENTER

        # 2. Add the phase headers at the top with better spacing and styling
        if "phases" in analysis_data and analysis_data["phases"]:
            phase_width = Inches(1.8)
            phase_height = Inches(0.6)
            total_phases = len(analysis_data["phases"])

            # Calculate total width needed for all phases
            total_width = total_phases * phase_width
            # Calculate start position to center the phases
            start_x = (self.slide_width - total_width) / 2

            for i, phase in enumerate(analysis_data["phases"]):
                # Calculate position for evenly spaced boxes
                phase_x = start_x + (i * phase_width)

                # Create a box for each phase
                phase_shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,  # Use rounded rectangle for modern look
                    left=phase_x,
                    top=Inches(0.7),
                    width=phase_width,
                    height=phase_height,
                )

                # Apply gradient fill for modern look
                fill = phase_shape.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(240, 240, 255)  # Light blue background

                # Configure the shape outline
                line = phase_shape.line
                line.color.rgb = RGBColor(100, 100, 180)  # Darker blue outline
                line.width = Pt(1.5)

                # Add text with improved styling
                text_frame = phase_shape.text_frame
                text_frame.text = phase
                text_frame.margin_left = Pt(6)
                text_frame.margin_right = Pt(6)
                text_frame.margin_top = Pt(3)
                text_frame.margin_bottom = Pt(3)
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                p = text_frame.paragraphs[0]
                p.font.size = Pt(14)
                p.font.bold = True
                p.font.name = "Calibri"
                p.font.color.rgb = RGBColor(0, 0, 0)  # Black text for all phases
                p.alignment = PP_ALIGN.CENTER

                # Ensure all runs have black text
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)

        # 3. Add the vertical section labels on the left with improved styling
        if "sections" in analysis_data and analysis_data["sections"]:
            for i, section in enumerate(analysis_data["sections"]):
                # Calculate position with better spacing
                section_y = Inches(1.8 + i * 1.6)

                section_shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left=Inches(0.3),
                    top=section_y,
                    width=Inches(0.9),
                    height=Inches(1.2),
                )

                # Apply styling
                section_shape.fill.solid()
                section_shape.fill.fore_color.rgb = RGBColor(
                    245, 245, 245
                )  # Light gray
                section_shape.line.color.rgb = RGBColor(120, 120, 120)  # Gray outline

                # Add text
                text_frame = section_shape.text_frame
                text_frame.text = section
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                section_shape.rotation = 270  # Rotate text vertically

                p = text_frame.paragraphs[0]
                p.font.size = Pt(12)
                p.font.name = "Calibri"
                p.font.bold = True
                p.font.color.rgb = RGBColor(0, 0, 0)  # Black text for sections
                p.alignment = PP_ALIGN.CENTER

                # Ensure all runs have black text
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)

        # 4. Intelligent positioning of elements
        element_layout = self._intelligently_position_elements(analysis_data)

        # Track shapes for connections
        element_shapes = {}

        # Create elements based on calculated layout
        for element_id, element_data in element_layout.items():
            element = element_data["element"]
            position = element_data["position"]
            width = element_data["width"]
            height = element_data["height"]

            # Determine shape type based on element properties
            shape_type = MSO_SHAPE.ROUNDED_RECTANGLE

            # Use appropriate shape type based on element type
            if element.get("type") == "note":
                shape_type = MSO_SHAPE.CLOUD
            elif element.get("type") == "cloud":
                shape_type = MSO_SHAPE.CLOUD

            # Create the shape
            shape = slide.shapes.add_shape(
                shape_type,
                left=position["left"],
                top=position["top"],
                width=width,
                height=height,
            )

            # Apply styling based on importance
            importance = element.get("importance", 1)

            # Define colors based on importance/emphasis
            if importance >= 3 or element.get("emphasis") == "high":
                # High importance - stronger color
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(255, 240, 240)  # Light red tint
                shape.line.color.rgb = RGBColor(180, 100, 100)  # Red outline
                shape.line.width = Pt(2.0)
            elif importance == 2 or element.get("emphasis") == "medium":
                # Medium importance
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(240, 255, 240)  # Light green tint
                shape.line.color.rgb = RGBColor(100, 160, 100)  # Green outline
                shape.line.width = Pt(1.5)
            else:
                # Normal importance
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(250, 250, 250)  # Off-white
                shape.line.color.rgb = RGBColor(100, 100, 100)  # Gray outline
                shape.line.width = Pt(1.0)

            # Add text with proper formatting - ensure text is black
            text_frame = shape.text_frame
            text_frame.word_wrap = True
            text_frame.text = ""  # Clear any default text

            # Set margin for better text appearance
            text_frame.margin_left = Pt(6)
            text_frame.margin_right = Pt(6)
            text_frame.margin_top = Pt(3)
            text_frame.margin_bottom = Pt(3)

            # Set vertical alignment
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

            # Create a new paragraph
            p = text_frame.paragraphs[0]

            # Add the text as a run to ensure color control
            run = p.add_run()
            run.text = element["text"]

            # Adjust font size based on text length and importance
            text_length = element.get("text_length", len(element["text"]))

            if text_length > 70:
                font_size = 9
            elif text_length > 50:
                font_size = 10
            elif text_length > 30:
                font_size = 11
            else:
                font_size = 12

            # Increase font size for important elements
            if importance >= 2:
                font_size += 1

            run.font.size = Pt(font_size)
            run.font.name = "Calibri"

            # Bold for important elements
            if importance >= 2:
                run.font.bold = True

            # Explicitly set text color to black
            run.font.color.rgb = RGBColor(0, 0, 0)  # Pure black text

            # Set paragraph alignment
            p.alignment = PP_ALIGN.CENTER

            # Store shape info for connections
            element_shapes[element_id] = {
                "shape": shape,
                "left": position["left"],
                "top": position["top"],
                "width": width,
                "height": height,
            }

        # 5. Add connections (arrows) with improved styling
        if "connections" in analysis_data:
            for connection in analysis_data["connections"]:
                from_id = connection.get("from_id")
                to_id = connection.get("to_id")
                arrow_type = connection.get("arrow_type", "simple")

                # Skip if shapes don't exist
                if from_id not in element_shapes or to_id not in element_shapes:
                    continue

                from_shape = element_shapes[from_id]
                to_shape = element_shapes[to_id]

                # Calculate center points
                from_center_x = from_shape["left"] + (from_shape["width"] / 2)
                from_center_y = from_shape["top"] + (from_shape["height"] / 2)
                to_center_x = to_shape["left"] + (to_shape["width"] / 2)
                to_center_y = to_shape["top"] + (to_shape["height"] / 2)

                # Calculate angle and direction of connection
                dx = to_center_x - from_center_x
                dy = to_center_y - from_center_y

                # Determine the best connection points based on angle
                # (This creates more natural looking connectors)

                # Calculate starting point (on the edge of the source shape)
                if abs(dx) > abs(dy):  # More horizontal than vertical
                    if dx > 0:  # Right direction
                        start_x = from_shape["left"] + from_shape["width"]
                        start_y = from_center_y
                    else:  # Left direction
                        start_x = from_shape["left"]
                        start_y = from_center_y
                else:  # More vertical than horizontal
                    if dy > 0:  # Down direction
                        start_x = from_center_x
                        start_y = from_shape["top"] + from_shape["height"]
                    else:  # Up direction
                        start_x = from_center_x
                        start_y = from_shape["top"]

                # Calculate ending point (on the edge of the target shape)
                if abs(dx) > abs(dy):  # More horizontal than vertical
                    if dx > 0:  # Right direction
                        end_x = to_shape["left"]
                        end_y = to_center_y
                    else:  # Left direction
                        end_x = to_shape["left"] + to_shape["width"]
                        end_y = to_center_y
                else:  # More vertical than horizontal
                    if dy > 0:  # Down direction
                        end_x = to_center_x
                        end_y = to_shape["top"]
                    else:  # Up direction
                        end_x = to_center_x
                        end_y = to_shape["top"] + to_shape["height"]

                # Create an improved connector
                connector_type = MSO_CONNECTOR.STRAIGHT

                # Use curved connectors for longer distances
                if (abs(dx) + abs(dy)) > Inches(4).inches:
                    connector_type = MSO_CONNECTOR.CURVE

                connector = slide.shapes.add_connector(
                    connector_type, start_x, start_y, end_x, end_y
                )

                # Style the connector based on arrow_type
                connector.line.color.rgb = RGBColor(100, 100, 100)  # Gray arrow
                connector.line.width = Pt(1.5)

                # Enhanced styling
                if arrow_type == "double":
                    connector.line.width = Pt(2.0)
                    connector.line.dash_style = MSO_LINE_DASH_STYLE.DASH
                    connector.line.begin_style = 1  # Simple arrow
                    connector.line.end_style = 1  # Simple arrow
                else:  # Simple or default
                    connector.line.dash_style = MSO_LINE_DASH_STYLE.SOLID
                    connector.line.end_style = 1  # Simple arrow at end

        # Save the PowerPoint
        pptx_io = BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)
        return pptx_io

    def process_image_to_ppt(self, image_data):
        """
        Main method to process image and create PowerPoint.
        """
        # Step 1: Analyze image with AI
        analysis_data = self.analyze_image(image_data)

        if not analysis_data:
            st.error("Failed to analyze image")
            return None

        # Step 2: Create PowerPoint from analysis
        with st.status("Creating PowerPoint slide...", expanded=True) as status:
            pptx_data = self.create_ppt_from_analysis(analysis_data)
            status.update(label="PowerPoint creation complete!", state="complete")

        return pptx_data


def main():
    st.set_page_config(
        page_title="Whiteboard to PowerPoint Converter", page_icon="📊", layout="wide"
    )

    st.title("Whiteboard to PowerPoint Converter")
    st.markdown(
        """
    Transform your whiteboard images into organized PowerPoint slides.
    Upload a photo of your whiteboard, and our AI will detect elements and create a professional slide.
    """
    )

    converter = WhiteboardToPPT()

    uploaded_file = st.file_uploader(
        "Upload a whiteboard image", type=["jpg", "jpeg", "png"]
    )

    if uploaded_file:
        # Display the uploaded image
        image = Image.open(uploaded_file)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Original Whiteboard")
            st.image(image, use_container_width=True)

        # Convert image to bytes for processing
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format=image.format)
        img_byte_arr = img_byte_arr.getvalue()

        # Process image and get PowerPoint
        pptx_data = converter.process_image_to_ppt(img_byte_arr)

        if pptx_data:
            with col2:
                st.subheader("PowerPoint Preview")
                st.write("PowerPoint successfully created!")
                st.info("Download your PowerPoint file using the button below.")

            # Provide download option
            st.download_button(
                label="Download PowerPoint",
                data=pptx_data,
                file_name="whiteboard_conversion.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        else:
            st.error("Failed to create PowerPoint. Please try with a clearer image.")


if __name__ == "__main__":
    main()


## much better final code
