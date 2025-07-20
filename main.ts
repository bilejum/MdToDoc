import { App, Editor, MarkdownView, Modal, Notice, Plugin, PluginSettingTab, Setting } from 'obsidian';

import { Document, Packer, Paragraph, TextRun } from "docx";


// Remember to rename these classes and interfaces!

interface MyPluginSettings {
	mySetting: string;
}

const DEFAULT_SETTINGS: MyPluginSettings = {
	mySetting: 'default'
}

export default class MyPlugin extends Plugin {
	settings: MyPluginSettings;

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
	async exportCurrentMarkdownToWord(markdownView: MarkdownView) {
		const file = markdownView.file;
		let content = '';
		if (markdownView.editor) {
			content = markdownView.editor.getValue();
		} else if (file) {
			content = await this.app.vault.read(file);
		}
		if (!content) {
			new Notice('无法获取当前文档内容');
			return;
		}

		// 简单的Markdown转docx实现（仅处理段落和标题）
		const lines = content.split(/\r?\n/);
		const docParagraphs: Paragraph[] = [];
		for (const line of lines) {
			if (/^# (.*)/.test(line)) {
				docParagraphs.push(new Paragraph({ text: line.replace(/^# /, ''), heading: 'Heading1' }));
			} else if (/^## (.*)/.test(line)) {
				docParagraphs.push(new Paragraph({ text: line.replace(/^## /, ''), heading: 'Heading2' }));
			} else if (/^### (.*)/.test(line)) {
				docParagraphs.push(new Paragraph({ text: line.replace(/^### /, ''), heading: 'Heading3' }));
			} else if (line.trim() === '') {
				docParagraphs.push(new Paragraph({ text: '' }));
			} else {
				docParagraphs.push(new Paragraph({ text: line }));
			}
		}

		const doc = new Document({
			sections: [
				{
					properties: {},
					children: docParagraphs
				}
			]
		});

		try {
			const blob = await Packer.toBlob(doc);
			const fileName = (file?.basename || '导出文档') + '.docx';
			const arrayBuffer = await blob.arrayBuffer();

			// 尝试在当前文件所在目录保存
			let savePath = fileName;
			if (file?.parent) {
				// 确保目录存在
				const exportDir = `${file.parent.path}/exports`;
				try {
					await this.app.vault.createFolder(exportDir);
				} catch (e) {
					// 如果目录已存在，忽略错误
				}
				savePath = `${exportDir}/${fileName}`;
			}

			await this.app.vault.createBinary(savePath, arrayBuffer);
			new Notice(`已导出到: ${savePath}`);
		} catch (error) {
			console.error('导出失败:', error);
			new Notice(`导出失败: ${error instanceof Error ? error.message : String(error)}`);
		}
	}

	async loadSettings() {
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
