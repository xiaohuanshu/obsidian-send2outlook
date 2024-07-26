import { App, Editor, MarkdownView, Notice, Plugin, PluginSettingTab, Setting } from 'obsidian';
import { marked } from 'marked';
import { exec } from 'child_process';

// Remember to rename these classes and interfaces!

interface Send2OutlookPluginSettings {
	// 默认收件人
	defaultRecipient: string;
	// 默认抄送
	defaultCc: string;
}

const DEFAULT_SETTINGS: Send2OutlookPluginSettings = {
	defaultRecipient: '',
	defaultCc: '',
}

export default class Send2OutlookPlugin extends Plugin {
	settings: Send2OutlookPluginSettings;

	setupEditorMenuEntry() {
		// 创建选项菜单
		this.registerEvent(this.app.workspace.on("file-menu", (menu, file, view) => {
			menu.addItem((item) => {
				item.setTitle("Send to Outlook").setIcon("clipboard-copy").onClick(async () => {
					this.sendEmail();
				});
			});
		}));
	}

	async onload() {
		await this.loadSettings();
		this.setupEditorMenuEntry();

		// This adds an editor command that can perform some operation on the current editor instance
		this.addCommand({
			id: 'send-to-outlook',
			name: 'Send To Outlook',
			editorCallback: (editor: Editor, view: MarkdownView) => {
				this.sendEmail();
			}
		});

		// This adds a settings tab so the user can configure various aspects of the plugin
		this.addSettingTab(new Send2OutlookPluginSettingTab(this.app, this));

	}

	onunload() {

	}

	sendEmail() {
		const noteFile = this.app.workspace.getActiveFile(); // Currently Open Note
		if (!noteFile) return; // Nothing Open
		const view = this.app.workspace.getActiveViewOfType(MarkdownView);
		// Make sure the user is editing a Markdown file.
		if (!view) return;
		// 获取笔记标题
		const title = noteFile.basename;
		// 获取笔记内容
		const text = view.editor.getDoc().getValue()
		// 渲染markdown
		const html = this.renderMarkdown(text);
		//发送
		this.callOutlook(title, html);
	}

	callOutlook(title: string, html: string) {
		const recipientEmail = this.settings.defaultRecipient;
		const ccEmail = this.settings.defaultCc;
		// 调用Applescript
		let script = `
		tell application "Microsoft Outlook"
			set emailSubject to "${title}"
			set emailContent to "${html.replace(/"/g, '\\"').replace(/\n/g, '\\n')}"
			set newEmail to make new outgoing message with properties {subject:emailSubject, content:emailContent}`;
		if (recipientEmail !== "") {
			const recipientEmails = recipientEmail.split(';');
			recipientEmails.forEach(email => {
				script += `
            make new recipient at newEmail with properties {email address:{address:"${email.trim()}"}}`;
			});
		}

		if (ccEmail !== "") {
			const ccEmails = ccEmail.split(';');
			ccEmails.forEach(email => {
				script += `
				make new cc recipient at newEmail with properties {email address:{address:"${email.trim()}"}}`;
			});
		}

		script += `
    		open newEmail
			activate
		end tell`;
		// 执行Applescript
		exec(`osascript -e '${script}'`, (err: any, stdout: any, stderr: any) => {
			if (err) {
				new Notice('Error: ' + err);
				return;
			}
			console.log(stdout);
		});
	}

	renderMarkdown(text: string) {
		return marked(text);
	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}

class Send2OutlookPluginSettingTab extends PluginSettingTab {
	plugin: Send2OutlookPlugin;

	constructor(app: App, plugin: Send2OutlookPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		new Setting(containerEl)
			.setName('默认收件人')
			.setDesc('使用Outlook发送邮件时默认的收件人')
			.addText(text => text
				.setPlaceholder('设置默认收件人')
				.setValue(this.plugin.settings.defaultRecipient)
				.onChange(async (value) => {
					this.plugin.settings.defaultRecipient = value;
					await this.plugin.saveSettings();
				}));
		new Setting(containerEl)
			.setName('默认抄送')
			.setDesc('使用Outlook发送邮件时默认的抄送人')
			.addText(text => text
				.setPlaceholder('设置默认抄送人')
				.setValue(this.plugin.settings.defaultCc)
				.onChange(async (value) => {
					this.plugin.settings.defaultCc = value;
					await this.plugin.saveSettings();
				}));
	}
}
