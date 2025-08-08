import { App, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile } from 'obsidian';

import { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, IImageOptions } from "docx";


// Remember to rename these classes and interfaces!

interface MyPluginSettings {
	mySetting: string;
	preserveCase: boolean;
	imageMaxWidth: number;
	imageMaxHeight: number;
}

const DEFAULT_SETTINGS: MyPluginSettings = {
	mySetting: 'default',
	preserveCase: true,
	imageMaxWidth: 400,
	imageMaxHeight: 300
}

export default class MyPlugin extends Plugin {
	public settings: MyPluginSettings;
	private currentMarkdownView: MarkdownView | null;
	private currentDoc: Document | null;

	constructor(app: App, manifest: any) {
		super(app, manifest);
		this.settings = Object.assign({}, DEFAULT_SETTINGS);
		this.currentMarkdownView = null;
		this.currentDoc = null;
	}

	private processFormattedText(text: string): TextRun[] {
		const children: TextRun[] = [];
		let currentIndex = 0;
		const boldRegex = /\*\*(.*?)\*\*|__(.*?)__/g;
		let match;

		while ((match = boldRegex.exec(text)) !== null) {
			if (match.index > currentIndex) {
				children.push(new TextRun({
					text: text.slice(currentIndex, match.index)
				}));
			}

			const boldText = match[1] || match[2];
			children.push(new TextRun({
				text: boldText,
				bold: true
			}));

			currentIndex = match.index + match[0].length;
		}

		if (currentIndex < text.length) {
			children.push(new TextRun({
				text: text.slice(currentIndex)
			}));
		}

		if (children.length === 0) {
			children.push(new TextRun({ text }));
		}

		return children;
	}

	private async processLine(line: string): Promise<Paragraph> {
		if (/^# (.*)/.test(line)) {
			const content = line.replace(/^# /, '');
			return new Paragraph({
				children: this.processFormattedText(content),
				heading: HeadingLevel.HEADING_1
			});
		}
		else if (/^## (.*)/.test(line)) {
			const content = line.replace(/^## /, '');
			return new Paragraph({
				children: this.processFormattedText(content),
				heading: HeadingLevel.HEADING_2
			});
		} else if (/^### (.*)/.test(line)) {
			const content = line.replace(/^### /, '');
			return new Paragraph({
				children: this.processFormattedText(content),
				heading: HeadingLevel.HEADING_3
			});
		} else if (line.trim() === '') {
			return new Paragraph({ text: '' });
		}
		else if (line.startsWith('> ') || line.startsWith('>')) {
			// 去掉 > 符号，仅保留内容
			return new Paragraph({
				children: this.processFormattedText(line.replace(/^> ?/, '')),
				indent: { left: 720 },
				spacing: { before: 240, after: 240 },
				shading: { type: 'solid', color: 'E8E8E8' }
			});
		}
		else if (line.startsWith('- ') || line.startsWith('* ')) {
			return new Paragraph({
				children: [new TextRun({ text: line.replace(/^[-*] /, '') })],
				bullet: {
					level: 0
				}
			});
		}
		else if (/^\d+\. /.test(line)) {
			return new Paragraph({
				children: [new TextRun({ text: line.replace(/^\d+\. /, '') })],
				numbering: {
					reference: "1",
					level: 0
				}
			});
		}
		else if (line.startsWith('```')) {
			return new Paragraph({ text: line.replace(/^```/, ''), style: 'codeBlock' });
		}
		else if (line.startsWith('`')) {
			return new Paragraph({ text: line.replace(/^`/, '').replace(/`$/, ''), style: 'inlineCode' });
		}
		else if (line.startsWith('---') || line.startsWith('***')) {
			return new Paragraph({ text: line.replace(/^[*-]{3,}$/, ''), style: 'horizontalLine' });
		}
		else if (line.startsWith('[') && line.includes('](')) {
			const match = line.match(/\[(.*?)\]\((.*?)\)/);
			if (match) {
				const [, text, url] = match;
				const hyperlink = new TextRun({ text: text, style: 'Hyperlink' });
				return new Paragraph({ children: [hyperlink] });
			}
		}
		else if (line.startsWith('![')) {
			const imageRun = await this.processImage(line);
			if (imageRun) {
				return new Paragraph({
					children: [imageRun],
					spacing: { before: 240, after: 240 }
				});
			}
			return new Paragraph({ text: `[图片处理失败: ${line}]` });
		}
		else if (line.startsWith('**') || line.startsWith('__')) {
			return new Paragraph({
				children: this.processFormattedText(line)
			});
		}
		else if (line.startsWith('*') || line.startsWith('_')) {
			return new Paragraph({
				children: [new TextRun({
					text: line.replace(/^\*|\_|\*$/g, ''),
					italics: true
				})]
			});
		}
		else if (line.startsWith('***') || line.startsWith('___')) {
			return new Paragraph({
				children: [new TextRun({
					text: line.replace(/^\*\*\*|\_\_\_|\*\*\*$/g, ''),
					bold: true,
					italics: true
				})]
			});
		}

		return new Paragraph({
			children: this.processFormattedText(line)
		});
	}

	private getImagePath(line: string): { url: string; alt: string } | null {
		// 处理 Obsidian 格式 ![[image.png]]
		const obsidianMatch = line.match(/!\[\[(.*?)\]\]/);
		if (obsidianMatch) {
			return {
				url: obsidianMatch[1],
				alt: obsidianMatch[1]
			};
		}

		// 处理标准 Markdown 格式 ![alt](url)
		const markdownMatch = line.match(/!\[(.*?)\]\((.*?)\)/);
		if (markdownMatch) {
			return {
				url: markdownMatch[2],
				alt: markdownMatch[1]
			};
		}

		return null;
	}

	private async processImage(line: string): Promise<ImageRun | null> {
		const imageInfo = this.getImagePath(line);
		if (!imageInfo) {
			return null;
		}

		try {
			const imageFile = this.app.vault.getAbstractFileByPath(imageInfo.url);
			if (!imageFile || !(imageFile instanceof TFile)) {
				console.error(`图片文件不存在: ${imageInfo.url}`);
				return null;
			}

			const imageData = await this.app.vault.readBinary(imageFile);
			const lowerUrl = imageInfo.url.toLowerCase();
			let imageType: 'png' | 'jpg' | 'gif' | 'bmp' = 'jpg';
			if (lowerUrl.endsWith('.png')) imageType = 'png';
			else if (lowerUrl.endsWith('.jpg') || lowerUrl.endsWith('.jpeg')) imageType = 'jpg';
			else if (lowerUrl.endsWith('.gif')) imageType = 'gif';
			else if (lowerUrl.endsWith('.bmp')) imageType = 'bmp';
			else if (lowerUrl.endsWith('.svg')) {
				new Notice('SVG 图片暂不支持导出为 Word。');
				return null;
			}

			// 获取图片的二进制数据前几个字节来判断图片尺寸
			let width = this.settings.imageMaxWidth;
			let height = this.settings.imageMaxHeight;

			// 尝试从图片数据中获取原始尺寸
			try {
				// 创建一个 Uint8Array 来读取图片头部信息
				const arr = new Uint8Array(imageData);
				let originalWidth = 0;
				let originalHeight = 0;

				// 检查是否是 PNG
				if (arr[0] === 0x89 && arr[1] === 0x50 && arr[2] === 0x4E && arr[3] === 0x47) {
					// PNG 图片尺寸在固定位置
					originalWidth = (arr[16] << 24) | (arr[17] << 16) | (arr[18] << 8) | arr[19];
					originalHeight = (arr[20] << 24) | (arr[21] << 16) | (arr[22] << 8) | arr[23];
				}
				// JPEG 的情况
				else if (arr[0] === 0xFF && arr[1] === 0xD8) {
					let pos = 2;
					while (pos < arr.length - 10) {
						if (arr[pos] === 0xFF && arr[pos + 1] === 0xC0) {
							originalHeight = (arr[pos + 5] << 8) | arr[pos + 6];
							originalWidth = (arr[pos + 7] << 8) | arr[pos + 8];
							break;
						}
						pos++;
					}
				}

				// 如果成功获取到原始尺寸，计算缩放后的尺寸
				if (originalWidth > 0 && originalHeight > 0) {
					const scale = Math.min(
						this.settings.imageMaxWidth / originalWidth,
						this.settings.imageMaxHeight / originalHeight,
						1
					);
					width = Math.round(originalWidth * scale);
					height = Math.round(originalHeight * scale);
					console.log(`图片 ${imageInfo.url} 原始尺寸: ${originalWidth}x${originalHeight}, 缩放后: ${width}x${height}`);
				}
			} catch (e) {
				console.log('获取图片尺寸失败，使用默认尺寸', e);
			}

			return new ImageRun({
				data: imageData,
				transformation: {
					width,
					height
				},
				type: imageType
			});
		} catch (error) {
			console.error('处理图片失败:', error);
			return null;
		}
	}

	async onload() {
		await this.loadSettings();

		// This creates an icon in the left ribbon.
		const ribbonIconEl = this.addRibbonIcon('file-down', '导出为word', (evt: MouseEvent) => {
			// 获取当前的 Markdown 视图
			const markdownView = this.app.workspace.getActiveViewOfType(MarkdownView);
			if (markdownView) {
				this.exportCurrentMarkdownToWord(markdownView);
			} else {
				new Notice('请在 Markdown 文件中使用此功能');
			}
		});
		// Perform additional things with the ribbon
		ribbonIconEl.addClass('my-plugin-ribbon-class');

		// This adds a simple command that can be triggered anywhere
		// This adds a settings tab so the user can configure various aspects of the plugin
		this.addSettingTab(new SampleSettingTab(this.app, this));

		// Add export to Word command (move to top, use English id/name)
		this.addCommand({
			id: 'export-current-markdown-to-word',
			name: '导出当前Markdown为Word Doc',
			checkCallback: (checking: boolean) => {
				const markdownView = this.app.workspace.getActiveViewOfType(MarkdownView);
				if (markdownView) {
					if (!checking) {
						this.exportCurrentMarkdownToWord(markdownView);
					}
					return true;
				}
				return false;
			}
		});

		// If the plugin hooks up any global DOM events (on parts of the app that doesn't belong to this plugin)
		// Using this function will automatically remove the event listener when this plugin is disabled.
		this.registerDomEvent(document, 'click', (evt: MouseEvent) => {
			console.log('click', evt);
		});

		// When registering intervals, this function will automatically clear the interval when the plugin is disabled.
		this.registerInterval(window.setInterval(() => console.log('setInterval'), 5 * 60 * 1000));
	}

	onunload() {

	}

	/**
	 * 将当前Markdown文档导出为Word（docx）
	 */
	private async getContent(markdownView: MarkdownView): Promise<string | null> {
		const file = markdownView.file;
		if (!file) {
			new Notice('当前没有打开的文件');
			return null;
		}

		let content = '';
		if (markdownView.editor) {
			content = markdownView.editor.getValue();
		} else {
			content = await this.app.vault.read(file);
		}

		if (!content) {
			new Notice('无法获取当前文档内容');
			return null;
		}

		return content;
	}

	async exportCurrentMarkdownToWord(markdownView: MarkdownView) {
		try {
			// 1. 准备文档内容
			this.currentMarkdownView = markdownView;
			const content = await this.getContent(markdownView);
			if (!content) return;

			// 2. 生成 Word 文档
			const lines = content.split(/\r?\n/);
			const paragraphs: Paragraph[] = [];
			for (const line of lines) {
				const paragraph = await this.processLine(line);
				paragraphs.push(paragraph);
			}
			const doc = new Document({
				numbering: {
					config: [
						{
							reference: "1",
							levels: [
								{
									level: 0,
									format: "decimal",
									text: "%1.",
									alignment: "start",
									style: {
										paragraph: {
											indent: { left: 720, hanging: 360 }
										}
									}
								}
							]
						}
					]
				},
				sections: [
					{
						properties: {},
						children: paragraphs
					}
				]
			});

			// 3. 转换为二进制数据
			const blob = await Packer.toBlob(doc);
			const arrayBuffer = await blob.arrayBuffer();

			// 4. 准备保存路径
			const fileName = (this.currentMarkdownView?.file?.basename || '导出文档') + '.docx';
			const currentDir = this.currentMarkdownView?.file?.parent;
			if (!currentDir) {
				new Notice('无法获取当前文件目录');
				return;
			}

			// 5. 创建导出目录
			const exportDir = `${currentDir.path}/exports`;
			try {
				const adapter = this.app.vault.adapter;
				if (!await adapter.exists(exportDir)) {
					await this.app.vault.createFolder(exportDir);
					// 等待目录创建完成
					await new Promise(res => setTimeout(res, 100));
				}
			} catch (e) {
				console.log('创建导出目录失败（可能已存在）:', e);
			}

			// 6. 生成不重复的文件名
			let finalFileName = fileName;
			let counter = 1;
			const adapter = this.app.vault.adapter;

			while (await adapter.exists(`${exportDir}/${finalFileName}`)) {
				const nameWithoutExt = fileName.slice(0, -5);
				finalFileName = `${nameWithoutExt}${counter}.docx`;
				counter++;
			}

			// 7. 保存文件
			const savePath = `${exportDir}/${finalFileName}`;
			try {
				await this.app.vault.createBinary(savePath, arrayBuffer);

				// 8. 验证文件是否创建成功
				if (await adapter.exists(savePath)) {
					new Notice(`已导出到: ${savePath}`);
				} else {
					throw new Error('文件创建后未找到');
				}
			} catch (error) {
				console.error('保存文件失败:', error);
				new Notice(`导出失败: ${error instanceof Error ? error.message : String(error)}`);
			}
		} catch (error) {
			// 只捕获未预期的错误
			console.error('未预期的错误:', error);
			new Notice(`导出失败: ${error instanceof Error ? error.message : String(error)}`);
		}
	} async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}


class SampleSettingTab extends PluginSettingTab {
	plugin: MyPlugin;

	constructor(app: App, plugin: MyPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		new Setting(containerEl)
			.setName('Setting #1')
			.setDesc('It\'s a secret')
			.addText(text => text
				.setPlaceholder('Enter your secret')
				.setValue(this.plugin.settings.mySetting)
				.onChange(async (value) => {
					this.plugin.settings.mySetting = value;
					await this.plugin.saveSettings();
				}));
	}
}
