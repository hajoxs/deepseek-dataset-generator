import sys
import json
import random
import time
import threading
from typing import List, Dict, Any
import requests
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QTextEdit, QSpinBox, QLineEdit, 
                            QPushButton, QProgressBar, QComboBox, QMessageBox, 
                            QFrame, QFileDialog, QSplitter)
from PyQt6.QtCore import Qt, pyqtSignal, QObject
from PyQt6.QtGui import QFont, QColor
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET
import oletools.oleobj
import re

class MathTypeProcessor:
    @staticmethod
    def extract_mathtype_equation(ole_object: bytes) -> str:
        """
        从OLE对象中提取MathType公式的文本表示
        """
        try:
            # 尝试解析OLE对象
            equation = ""
            ole_parser = oletools.oleobj.OleObject(ole_object)
            if ole_parser.class_name == "Equation.3":  # MathType equation
                # 获取公式数据
                equation_data = ole_parser.get_stream('Equation Native')
                if equation_data:
                    # 将二进制数据转换为文本格式
                    # 这里使用简单的ASCII转换，实际应用中可能需要更复杂的解析
                    equation = equation_data.decode('ascii', errors='ignore')
                    # 清理和格式化公式文本
                    equation = MathTypeProcessor.clean_equation_text(equation)
            return equation
        except Exception as e:
            print(f"Error extracting MathType equation: {e}")
            return ""

    @staticmethod
    def clean_equation_text(equation: str) -> str:
        """
        清理和格式化公式文本
        """
        # 移除非打印字符
        equation = ''.join(char for char in equation if char.isprintable())
        # 移除多余空格
        equation = ' '.join(equation.split())
        # 添加LaTeX风格的数学标记
        if equation and not equation.startswith('$'):
            equation = f'${equation}$'
        return equation
    @staticmethod
    def process_paragraph(paragraph) -> str:
        """
        处理文档段落，提取文本和公式
        """
        text_parts = []
        
        for run in paragraph.runs:
            # 处理普通文本
            if run.text:
                text_parts.append(run.text)
            
            # 处理嵌入的对象
            for element in run._element.findall('.//w:object', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                ole_objects = element.findall('.//o:OLEObject', {'o': 'urn:schemas-microsoft-com:office:office'})
                for ole_obj in ole_objects:
                    if ole_obj.get('ProgID', '').startswith('Equation'):
                        # 获取OLE对象数据
                        ole_data = ole_obj.get('data', '')
                        if ole_data:
                            equation = MathTypeProcessor.extract_mathtype_equation(ole_data)
                            if equation:
                                text_parts.append(equation)
        
        return ' '.join(text_parts)

class DocumentProcessor:
    @staticmethod
    def process_docx(file_path: str) -> str:
        """
        处理Word文档，提取文本和公式
        """
        doc = docx.Document(file_path)
        content_parts = []
        
        for paragraph in doc.paragraphs:
            content = MathTypeProcessor.process_paragraph(paragraph)
            if content.strip():
                content_parts.append(content)
        
        return '\n'.join(content_parts)

class WorkerSignals(QObject):
    progress = pyqtSignal(int)
    result = pyqtSignal(str)
    error = pyqtSignal(str)
    finished = pyqtSignal()

class DeepSeekDatasetGenerator:
    def __init__(self, base_url: str = "http://localhost:11434/api/generate"):
        self.base_url = base_url
        self.model = "llama2"
        
    def set_model(self, model: str):
        self.model = model
    
    def set_base_url(self, base_url: str):
        self.base_url = base_url
        
    def _filter_thinking_content(self, text: str) -> str:
        """
        过滤掉回复中的思考过程，保留实际回答内容
        """
        # 移除<think>...</think>格式
        text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL | re.IGNORECASE)
        
        # 移除Let me think格式
        text = re.sub(r'Let me think.*?\n', '', text, flags=re.IGNORECASE)
        
        # 移除Thinking:格式
        text = re.sub(r'Thinking:.*?\n', '', text, flags=re.IGNORECASE)
        
        # 移除[thinking]...[/thinking]格式
        text = re.sub(r'\[thinking\].*?\[/thinking\]', '', text, flags=re.DOTALL | re.IGNORECASE)
        
        # 移除多余的空行
        text = re.sub(r'\n\s*\n', '\n', text)
        
        return text.strip()
        
    def generate_questions(self, context: str, num_questions: int = 3) -> List[str]:
        prompt = f"""
        基于以下文本生成{num_questions}个相关的问题：
        
        {context}
        
        请直接列出问题，每行一个，不要加序号或其他标记。
        """
        
        response = self._call_ollama(prompt)
        # 过滤思考内容
        response = self._filter_thinking_content(response)
        questions = [q.strip() for q in response.split('\n') if q.strip()]
        return questions[:num_questions]
    
    def generate_answer(self, context: str, question: str) -> str:
        prompt = f"""
        基于以下文本回答问题：
        
        文本内容：
        {context}
        
        问题：{question}
        
        请提供详细且准确的回答。
        """
        
        response = self._call_ollama(prompt)
        # 过滤思考内容
        return self._filter_thinking_content(response)
    
    def _call_ollama(self, prompt: str) -> str:
        data = {
            "model": self.model,
            "prompt": prompt,
            "stream": False
        }
        
        try:
            response = requests.post(self.base_url, json=data)
            response.raise_for_status()
            return response.json()["response"]
        except Exception as e:
            raise Exception(f"调用Ollama API时出错: {e}")

class StyleHelper:
    @staticmethod
    def get_light_style():
        return """
        QMainWindow {
            background-color: #f8f9fa;
        }
        QWidget {
            color: #212529;
            font-size: 14px;
        }
        QTextEdit {
            background-color: #ffffff;
            border: 1px solid #dee2e6;
            border-radius: 5px;
            padding: 8px;
            color: #212529;
        }
        QPushButton {
            background-color: #0d6efd;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 5px;
            font-weight: bold;
            min-width: 100px;
        }
        QPushButton:hover {
            background-color: #0b5ed7;
        }
        QPushButton:disabled {
            background-color: #6c757d;
        }
        QLineEdit, QSpinBox, QComboBox {
            background-color: #ffffff;
            border: 1px solid #dee2e6;
            border-radius: 5px;
            padding: 8px;
            color: #212529;
            min-height: 20px;
        }
        QProgressBar {
            border: 1px solid #dee2e6;
            border-radius: 5px;
            text-align: center;
            background-color: #ffffff;
            min-height: 20px;
        }
        QProgressBar::chunk {
            background-color: #0d6efd;
            border-radius: 5px;
        }
        QLabel {
            color: #212529;
            font-weight: bold;
        }
        QFrame {
            border: 1px solid #dee2e6;
            border-radius: 5px;
            background-color: #ffffff;
            padding: 10px;
        }
        QSplitter::handle {
            background-color: #dee2e6;
        }
        """

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.generator = DeepSeekDatasetGenerator()
        self.signals = WorkerSignals()
        self.dataset = None
        self.import_btn = None  # Add this line
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle('DeepSeek数据集生成器')
        self.setMinimumSize(1425, 950)
        self.setStyleSheet(StyleHelper.get_light_style())
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 主布局
        layout = QHBoxLayout(main_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 创建左右分割器
        splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(splitter)
        
        # 左侧面板（输入和输出）
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setSpacing(20)
        
        # 输入区域
        input_group = QFrame()
        input_layout = QVBoxLayout(input_group)
        
        input_header = QHBoxLayout()
        input_label = QLabel("输入文本:")
        self.char_count_label = QLabel("0/1000")  # 新增字符计数标签
        import_btn = QPushButton("导入文件")
        import_btn.clicked.connect(self.import_file)
        self.import_btn = import_btn  # Add this line
        input_header.addWidget(input_label)
        input_header.addWidget(self.char_count_label)  # 新增
        input_header.addStretch()
        input_header.addWidget(import_btn)
        
        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("在此输入需要生成问答对的文本内容（1000字符以内），或点击“导入文件”按钮导入txt文件或docx文件...")
        self.text_input.textChanged.connect(self._text_input_text_changed)

        input_layout.addLayout(input_header)
        input_layout.addWidget(self.text_input)
        
        # 输出区域
        output_group = QFrame()
        output_layout = QVBoxLayout(output_group)
        output_label = QLabel("生成结果:")
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(False)  # 允许编辑
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_text)
        
        left_layout.addWidget(input_group, stretch=1)
        left_layout.addWidget(output_group, stretch=1)
        
        # 右侧面板（设置和控制）
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setSpacing(20)
        
        # 参数设置区域
        params_group = QFrame()
        params_layout = QVBoxLayout(params_group)
        params_layout.setSpacing(15)
        
        # 模型选择
        model_label = QLabel("选择模型:")
        self.model_combo = QComboBox()
        self.model_combo.addItems(["deepseek-r1:1.5b","deepseek-r1:7b", "deepseek-r1:8b",
                                   "deepseek-r1:14b","deepseek-r1:32b","deepseek-r1:70b",
                                   "自定义"])
        # 添加自定义模型名称输入组件
        self.custom_model_label = QLabel("自定义模型名称:")
        self.custom_model_input = QLineEdit()
        self.custom_model_label.hide()
        self.custom_model_input.hide()
        
        # 信号连接
        self.model_combo.currentIndexChanged.connect(self.on_model_selection_change)
        
        model_url_label = QLabel("API地址:")
        self.model_url_input = QLineEdit("http://localhost:11434/api/generate")
        
        # 问答对数量
        pairs_label = QLabel("问答对数量:（1-500）")
        self.num_pairs = QSpinBox()
        self.num_pairs.setRange(1, 500)
        self.num_pairs.setValue(3)
        
        # 文件名设置
        file_label = QLabel("输出文件:")
        self.file_input = QLineEdit("deepseek_dataset.json")
        
        params_layout.addWidget(model_label)
        params_layout.addWidget(self.model_combo)
        params_layout.addWidget(self.custom_model_label)  
        params_layout.addWidget(self.custom_model_input)  
        params_layout.addWidget(model_url_label)
        params_layout.addWidget(self.model_url_input)
        params_layout.addWidget(pairs_label)
        params_layout.addWidget(self.num_pairs)
        params_layout.addWidget(file_label)
        params_layout.addWidget(self.file_input)
        
        # 进度条
        self.progress = QProgressBar()
        
        # 控制按钮
        button_group = QFrame()
        button_layout = QVBoxLayout(button_group)
        button_layout.setSpacing(10)
        
        self.generate_btn = QPushButton("生成数据集")
        self.generate_btn.clicked.connect(self.generate_dataset)
        
        self.save_btn = QPushButton("保存结果")
        self.save_btn.clicked.connect(self.save_dataset)
        self.save_btn.setEnabled(False)
        
        button_layout.addWidget(self.generate_btn)
        button_layout.addWidget(self.save_btn)
        
        # 格式选择
        format_label = QLabel("输出格式:")
        self.format_combo = QComboBox()
        self.format_combo.addItems(["对话格式", "JSON Lines"])
        params_layout.addWidget(format_label)
        params_layout.addWidget(self.format_combo)
        
        right_layout.addWidget(params_group)
        right_layout.addWidget(self.progress)
        right_layout.addWidget(button_group)
        right_layout.addStretch()
        
        # 添加面板到分割器
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 2)  # 左侧占比更大
        splitter.setStretchFactor(1, 1)
        
        # 信号连接
        self.signals.progress.connect(self.update_progress)
        self.signals.result.connect(self.update_output)
        self.signals.error.connect(self.show_error)
        self.signals.finished.connect(self.generation_finished)
        
        # 在按钮区域添加合并按钮
        self.merge_btn = QPushButton("合并JSON文件")
        self.merge_btn.clicked.connect(self.merge_json_files)
        button_layout.addWidget(self.merge_btn)
        
        # 添加退出按钮
        self.exit_btn = QPushButton("退出")
        self.exit_btn.clicked.connect(self.close)
        button_layout.addWidget(self.exit_btn)
        
    def _text_input_text_changed(self):
        current_text = self.text_input.toPlainText()
        text_length = len(current_text)
        max_length = 1000
        
        # 更新字符计数标签
        self.char_count_label.setText(f"{min(text_length, max_length)}/{max_length}")
        
        if text_length > max_length:
            self.text_input.setText(current_text[:max_length])
            QMessageBox.warning(self, "警告", "输入的文本长度不能超过1000字！")
        
    def on_model_selection_change(self, index):
        """处理模型选择变化事件"""
        selected = self.model_combo.currentText()
        if selected == "自定义":
            self.custom_model_label.show()
            self.custom_model_input.show()
        else:
            self.custom_model_label.hide()
            self.custom_model_input.hide()
            
    def get_selected_model(self):
        """获取当前选择的模型"""
        selected = self.model_combo.currentText()
        if selected == "自定义":
            custom_name = self.custom_model_input.text().strip()
            return custom_name if custom_name else "custom-model"
        return selected 
    
    def import_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "导入文件",
            "",
            "Word Files (*.docx);;Text Files (*.txt);;All Files (*.*)"
        )
        
        if file_name:
            try:
                if file_name.endswith('.docx'):
                    # 使用增强的文档处理器处理Word文档
                    text = DocumentProcessor.process_docx(file_name)
                else:
                    with open(file_name, 'r', encoding='utf-8') as f:
                        text = f.read()
                
                self.text_input.setText(text)
                
                 # 手动触发字符数更新
                self._text_input_text_changed()
                
                # 限制导入的文本为1000字以内
                max_length = 1000
                if len(text) > max_length:
                    text = text[:max_length]
                    QMessageBox.warning(self, "警告", "导入的文件内容因长度限制已截断为1000字！")
                    
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导入文件时出错：{str(e)}")

    
    def generate_dataset(self):
        context = self.text_input.toPlainText().strip()
        if not context:
            QMessageBox.warning(self, "警告", "请输入文本内容！")
            return
        
        # 禁用相关控件
        self.model_combo.setEnabled(False)
        self.model_url_input.setEnabled(False)
        self.num_pairs.setEnabled(False)
        self.generate_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.import_btn.setEnabled(False) 
        self.progress.setValue(0)
        self.output_text.clear()
            
        self.generate_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.progress.setValue(0)
        self.output_text.clear()
        
        # 更新选择的模型
        self.generator.set_model(self.get_selected_model())
        model_url = self.model_url_input.text().strip()
        if model_url:
            self.generator.set_base_url(model_url)
        
        # 在新线程中运行生成过程
        def generate():
            try:
                dataset = []
                num_pairs = self.num_pairs.value()
                questions = self.generator.generate_questions(context, num_pairs)
            
                for i, question in enumerate(questions):
                    answer = self.generator.generate_answer(context, question)
                
                    qa_pair = {
                        "conversations": [
                            {
                                "role": "user",
                                "content": question
                            },
                            {
                                "role": "assistant",
                                "content": answer
                            }
                        ]
                    }
                    dataset.append(qa_pair)
                
                    # 更新进度
                    progress = int((i + 1) / num_pairs * 100)
                    self.signals.progress.emit(progress)
                
                    # 更新输出
                    self.dataset = dataset
                    self.signals.result.emit(json.dumps(dataset, ensure_ascii=False, indent=2))
                
                    time.sleep(random.uniform(0.5, 1))  # 减少等待时间
            
                self.signals.finished.emit()
            
            except Exception as e:
                self.signals.error.emit(str(e))
                self.signals.finished.emit()
    
        threading.Thread(target=generate, daemon=True).start()
    
    def update_progress(self, value):
        self.progress.setValue(value)
        
    def convert_display_to_json(self, display_text: str) -> List[Dict]:
        """
        将显示格式转换回JSON格式，正确处理多行内容
        """
        dataset = []
        sections = display_text.split('-' * 50)
        
        for section in sections:
            if not section.strip():
                continue
                
            # 初始化当前问答对
            current_qa = {'conversations': []}
            
            # 分割问题和答案部分
            parts = section.split('答案：')
            if len(parts) != 2:
                continue
                
            question_part = parts[0]
            answer_part = parts[1]
            
            # 提取问题内容（移除"问题 N："标记）
            question_content = question_part.split('：', 1)[-1].strip()
            # 提取答案内容
            answer_content = answer_part.strip()
            
            # 只有当问题和答案都不为空时才添加到数据集
            if question_content and answer_content:
                current_qa['conversations'] = [
                    {
                        'role': 'user',
                        'content': question_content
                    },
                    {
                        'role': 'assistant',
                        'content': answer_content
                    }
                ]
                dataset.append(current_qa)
        
        return dataset

    def format_qa_pairs(self, dataset: List[Dict]) -> str:
        """
        将数据集格式化为易读的问答对显示格式，保持原始格式
        """
        formatted_text = []
        for i, qa_pair in enumerate(dataset, 1):
            conversations = qa_pair.get('conversations', [])
            if len(conversations) >= 2:
                question = conversations[0].get('content', '')
                answer = conversations[1].get('content', '')
                
                formatted_text.append(f"问题 {i}：\n{question}\n")
                formatted_text.append(f"答案：\n{answer}\n")
                formatted_text.append("-" * 50)
        
        return "\n".join(formatted_text)

    def update_output(self, text: str):
        """
        更新输出显示，保持原始JSON数据
        """
        try:
            # 将JSON字符串转换为Python对象
            dataset = json.loads(text)
            # 格式化为易读的显示格式
            formatted_text = self.format_qa_pairs(dataset)
            self.output_text.setText(formatted_text)
            # 保存原始数据集
            self.dataset = dataset
        except Exception as e:
            QMessageBox.critical(self, "错误", f"格式化输出时出错：{str(e)}")
    
    def show_error(self, message):
        QMessageBox.critical(self, "错误", message)

    def generation_finished(self):
        self.model_combo.setEnabled(True)
        self.model_url_input.setEnabled(True)
        self.num_pairs.setEnabled(True)
        self.generate_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        self.import_btn.setEnabled(True)  # Add this line
            # 修改save_dataset方法：
    def save_dataset(self):
        if not self.dataset:
            return
            
        try:
            # 获取选择的格式
            selected_format = self.format_combo.currentText()
            
            # 转换数据格式
            if selected_format == "JSON Lines":
                converted_data = []
                for item in self.dataset:
                    question = item["conversations"][0]["content"]
                    answer = item["conversations"][1]["content"]
                    converted_data.append({"text": f"{question} {answer}"})
            else:  # 对话格式
                converted_data = self.dataset

            # 设置文件过滤器和默认扩展名
            file_filters = {
                "对话格式": "JSON Files (*.json)",
                "JSON Lines": "JSON Lines Files (*.jsonl)"
            }
            default_ext = ".json" if selected_format == "对话格式" else ".jsonl"

            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "保存数据集",
                self.file_input.text().rsplit('.', 1)[0] + default_ext,
                file_filters[selected_format]
            )
            
            if file_name:
                with open(file_name, 'w', encoding='utf-8') as f:
                    if selected_format == "JSON Lines":
                        for item in converted_data:
                            f.write(json.dumps(item, ensure_ascii=False) + '\n')
                    else:
                        json.dump(converted_data, f, ensure_ascii=False, indent=2)
                        
                QMessageBox.information(self, "成功", f"数据集已保存到 {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存文件时出错：{str(e)}")

    # 修改merge_json_files方法以支持不同格式：
    def merge_json_files(self):
        """合并多个JSON文件，并按照当前选择的格式保存"""
        file_names, _ = QFileDialog.getOpenFileNames(
            self,
            "选择要合并的JSON文件",
            "",
            "JSON Files (*.json *.jsonl)"
        )
        if not file_names:
            return

        # 获取当前选择的格式
        selected_format = self.format_combo.currentText()
        is_jsonl = selected_format == "JSON Lines"

        try:
            # 合并数据（统一转换为对话格式）
            merged_data = []
            for file_name in file_names:
                with open(file_name, 'r', encoding='utf-8') as f:
                    if file_name.endswith('.jsonl'):
                        data = [json.loads(line) for line in f]
                    else:
                        data = json.load(f)

                    for item in data:
                        if "text" in item:  # 转换JSON Lines为对话格式
                            text = item["text"].split(' ', 1)
                            if len(text) >= 2:
                                merged_data.append({
                                    "conversations": [
                                        {"role": "user", "content": text[0]},
                                        {"role": "assistant", "content": text[1]}
                                    ]
                                })
                        elif "conversations" in item:  # 保持对话格式
                            merged_data.append(item)

            # 设置保存参数
            default_ext = ".jsonl" if is_jsonl else ".json"
            file_filter = "JSON Lines Files (*.jsonl)" if is_jsonl else "JSON Files (*.json)"
            
            save_name, _ = QFileDialog.getSaveFileName(
                self,
                "保存合并后的文件",
                f"merged_dataset{default_ext}",
                file_filter
            )

            if save_name:
                # 转换为目标格式
                if is_jsonl:
                    output_data = [
                        {"text": f"{item['conversations'][0]['content']} {item['conversations'][1]['content']}"}
                        for item in merged_data
                    ]
                else:
                    output_data = merged_data

                # 写入文件
                with open(save_name, 'w', encoding='utf-8') as f:
                    if is_jsonl:
                        for item in output_data:
                            f.write(json.dumps(item, ensure_ascii=False) + '\n')
                    else:
                        json.dump(output_data, f, ensure_ascii=False, indent=2)
                
                QMessageBox.information(self, "成功", f"文件已保存到 {save_name}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理文件时出错：{str(e)}")

def main():
    app = QApplication(sys.argv)
    
    # 设置应用程序级别的字体
    # font = QFont("Microsoft YaHei", 10)  # 使用微软雅黑字体
    # font = QFont("SimSun", 10)
    font = QFont("kaiTi", 10) # 设置字体为楷体
    
    # GUI上切换字体
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

