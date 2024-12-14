from pptx import Presentation
from pptx.util import Inches, Pt
import tempfile
import re
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import numpy as np

def extract_metrics(text):
    """Extract numerical data and metrics from analysis text."""
    metrics = {
        'values': [],
        'labels': [],
        'percentages': []
    }
    
    # Extract percentages
    percentage_matches = re.findall(r'(\d+(?:\.\d+)?)\s*%\s*(?:of|in|for)?\s*([^,.]*)', text)
    for value, label in percentage_matches:
        metrics['values'].append(float(value))
        metrics['labels'].append(label.strip() or f"Metric {len(metrics['labels'])+1}")
        metrics['percentages'].append(True)
    
    # Extract other numerical values
    number_matches = re.findall(r'([A-Za-z\s]+):\s*(\d+(?:\.\d+)?)', text)
    for label, value in number_matches:
        metrics['values'].append(float(value))
        metrics['labels'].append(label.strip())
        metrics['percentages'].append(False)
    
    return metrics

def add_chart_slide(prs, metrics, chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED):
    """Add a slide with the specified chart type."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    # Add chart
    chart_data = CategoryChartData()
    chart_data.categories = metrics['labels']
    chart_data.add_series('Values', metrics['values'])
    
    x, y, cx, cy = Inches(1), Inches(1), Inches(8), Inches(5.5)
    chart = slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data).chart
    
    # Customize chart
    chart.has_legend = True
    chart.has_title = True
    chart.chart_title.text_frame.text = "Performance Metrics"

def create_presentation(analysis, template_data=None):
    """Create presentation with enhanced analytics"""
    prs = Presentation()
    
    # Title slide
    _add_title_slide(prs, template_data)
    
    # Executive Summary
    _add_summary_slide(prs, analysis, template_data)
    
    # Key Metrics Dashboard
    metrics = extract_metrics(analysis)
    if metrics['values']:
        _add_metrics_dashboard(prs, metrics, template_data)
    
    # Detailed Analysis Slides
    sections = _structure_content(analysis)
    for section in sections:
        if 'performance' in section.lower() or 'analysis' in section.lower():
            _add_analysis_slide(prs, section, template_data)
    
    # Trend Analysis
    _add_trend_analysis(prs, analysis, metrics, template_data)
    
    # Segment Analysis
    _add_segment_analysis(prs, analysis, template_data)
    
    # Recommendations
    _add_recommendations_slide(prs, analysis, template_data)
    
    # Save presentation
    temp_ppt = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
    prs.save(temp_ppt.name)
    return temp_ppt.name

def _structure_content(analysis):
    """Better structure the content for presentation"""
    sections = []
    current_section = []
    
    for line in analysis.split('\n'):
        if line.strip():
            if any(key in line.lower() for key in ['summary:', 'overview:', 'metrics:', 'conclusion:']):
                if current_section:
                    sections.append('\n'.join(current_section))
                current_section = [line]
            else:
                current_section.append(line)
    
    if current_section:
        sections.append('\n'.join(current_section))
    
    return sections

def _add_content_slide(prs, content, styles, layouts):
    """Add content slide with template styling"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    
    # Add title
    if hasattr(slide.shapes, "title"):
        title = slide.shapes.title
        first_line = content.split('\n')[0]
        title.text = first_line
        
        # Apply title styling
        if 'content_title' in styles:
            title_style = styles['content_title']
            if hasattr(title.text_frame.paragraphs[0], 'font'):
                font = title.text_frame.paragraphs[0].font
                font.name = title_style.get('font_name', 'Calibri')
                font.size = title_style.get('size', Pt(32))
                font.bold = title_style.get('bold', True)

    # Add content
    if hasattr(slide.shapes, "placeholders"):
        body_shape = slide.shapes.placeholders[1]
        if body_shape.has_text_frame:
            tf = body_shape.text_frame
            tf.text = '\n'.join(content.split('\n')[1:])

def _apply_template_style(shape, style):
    """Apply template styling to a shape"""
    if hasattr(shape, "text_frame"):
        for paragraph in shape.text_frame.paragraphs:
            if hasattr(paragraph, "font"):
                font = paragraph.font
                font.name = style.get('font_name', 'Calibri')
                font.size = style.get('size', Pt(32))
                font.bold = style.get('bold', True)

def _add_summary_slide(prs, analysis, template_data):
    """Add summary slide matching template style"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Executive Summary"
    
    # Extract summary section
    summary = re.search(r'summary:?(.+?)(?=\n\n|\Z)', analysis, re.I | re.S)
    if summary:
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        tf.text = summary.group(1).strip()
        
        # Apply template bullet patterns if available
        if template_data and 'structure' in template_data:
            bullet_patterns = template_data['structure'].get('bullet_patterns', [])
            if bullet_patterns:
                for paragraph in tf.paragraphs[1:]:
                    paragraph.level = 1

def _add_chart_slides(prs, metrics, template_data=None):
    """Add chart slides based on metrics and template preferences."""
    if not metrics['values']:
        return

    # Add column chart for general metrics
    add_chart_slide(prs, metrics)
    
    # If we have percentage metrics, create a pie chart
    percentage_metrics = {
        'values': [v for v, is_pct in zip(metrics['values'], metrics['percentages']) if is_pct],
        'labels': [l for l, is_pct in zip(metrics['labels'], metrics['percentages']) if is_pct],
        'percentages': [True] * len([p for p in metrics['percentages'] if p])
    }
    
    if percentage_metrics['values']:
        add_chart_slide(prs, percentage_metrics, XL_CHART_TYPE.PIE)

def _add_metrics_dashboard(prs, metrics, template_data):
    """Add a dashboard-style slide with multiple charts"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    title = slide.shapes.title
    title.text = "Key Performance Metrics"
    
    # Split metrics for different charts
    perf_metrics = {k: v for k, v in zip(metrics['labels'], metrics['values']) 
                   if 'performance' in k.lower() or 'growth' in k.lower()}
    
    # Add multiple small charts
    if perf_metrics:
        _add_small_chart(slide, perf_metrics, XL_CHART_TYPE.COLUMN_CLUSTERED, 
                        Inches(0.5), Inches(1.5), Inches(4.5), Inches(3.5))
    
    # Add pie chart for percentages
    percentage_metrics = {
        'values': [v for v, p in zip(metrics['values'], metrics['percentages']) if p],
        'labels': [l for l, p in zip(metrics['labels'], metrics['percentages']) if p]
    }
    if percentage_metrics['values']:
        _add_small_chart(slide, percentage_metrics, XL_CHART_TYPE.PIE,
                        Inches(5.5), Inches(1.5), Inches(4.5), Inches(3.5))

def _add_trend_analysis(prs, analysis, metrics, template_data):
    """Add trend analysis with line charts"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    title = slide.shapes.title
    title.text = "Trend Analysis"
    
    # Extract trend data
    trend_matches = re.findall(r'trend:?\s*([^:]+):\s*(\d+(?:\.\d+)?)', analysis, re.I)
    if trend_matches:
        trend_data = {label.strip(): float(value) for label, value in trend_matches}
        _add_small_chart(slide, trend_data, XL_CHART_TYPE.LINE,
                        Inches(1), Inches(1.5), Inches(8), Inches(5))

def _add_small_chart(slide, data, chart_type, x, y, cx, cy):
    """Add a chart to the slide at specified position"""
    if not data:
        return  # Ensure data is not empty
    
    chart_data = CategoryChartData()
    chart_data.categories = list(data.keys())
    chart_data.add_series('Values', list(data.values()))
    
    chart = slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data).chart
    chart.has_legend = True

def _add_recommendations_slide(prs, analysis, template_data):
    """Add recommendations with impact analysis"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Recommendations & Impact Analysis"
    
    # Extract recommendations and their impacts
    recommendations = re.findall(r'recommend.*?:\s*([^(]+)\((\d+(?:\.\d+)?%?)\)', analysis, re.I)
    
    body_shape = slide.shapes.placeholders[1]
    tf = body_shape.text_frame
    
    for rec, impact in recommendations:
        p = tf.add_paragraph()
        p.text = f"• {rec.strip()} (Impact: {impact})"
        p.level = 0

def _add_title_slide(prs, template_data=None):
    """Create title slide matching reference template style"""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1] if len(slide.placeholders) > 1 else None
    
    # Apply template styling if available
    if template_data and 'styles' in template_data:
        title_style = template_data['styles'].get('title', {})
        _apply_template_style(title, title_style)
        
        # Apply background color if specified in template
        if 'color_scheme' in template_data and template_data['color_scheme']:
            try:
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = template_data['color_scheme'][0]
            except:
                pass

    title.text = "Document Analysis Report"
    if subtitle:
        subtitle.text = "AI-Powered Analysis"

def _add_analysis_slide(prs, section, template_data):
    """Add analysis slide matching reference template format"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])  # Using two-content layout
    
    # Extract section title and content
    title = slide.shapes.title
    section_lines = section.split('\n')
    title.text = section_lines[0]
    
    # Extract metrics and key points
    metrics = []
    key_points = []
    
    for line in section_lines[1:]:
        if ':' in line and any(char.isdigit() for char in line):
            metrics.append(line.strip())
        elif line.strip():
            key_points.append(line.strip())
    
    # Add content to left and right placeholders
    if len(slide.placeholders) > 2:
        left_content = slide.placeholders[1]
        right_content = slide.placeholders[2]
        
        # Add metrics to left side
        if metrics:
            tf = left_content.text_frame
            tf.text = "Key Metrics"
            for metric in metrics:
                p = tf.add_paragraph()
                p.text = f"• {metric}"
                p.level = 1
        
        # Add key points to right side
        if key_points:
            tf = right_content.text_frame
            tf.text = "Analysis"
            for point in key_points:
                p = tf.add_paragraph()
                p.text = f"• {point}"
                p.level = 1

    # Apply template styling
    if template_data and 'styles' in template_data:
        content_style = template_data['styles'].get('body', {})
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                _apply_template_style(shape, content_style)

def _add_segment_analysis(prs, analysis, template_data):
    """Add segment analysis slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "Segment Analysis"
    
    # Extract segment data
    segments = re.findall(r'segment:?\s*([^:]+):\s*(\d+(?:\.\d+)?)', analysis, re.I)
    if segments:
        segment_data = {label.strip(): float(value) for label, value in segments}
        _add_small_chart(slide, segment_data, XL_CHART_TYPE.BAR_CLUSTERED,
                        Inches(1), Inches(1.5), Inches(8), Inches(5))
