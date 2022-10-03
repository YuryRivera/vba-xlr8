import { commands, ExtensionContext, window, workspace } from "vscode";
import * as child from 'child_process';
import * as path from 'path';


const vbaRunningContext = 'vbaxlr8run';
let proccess: child.ChildProcess | undefined = undefined;

export function activate(context: ExtensionContext) {
	const folders = workspace.workspaceFolders;
	if(!folders || folders.length < 1){
		window.showErrorMessage("VBA XLR8 only works for valid project folder, please open the folder containing the vba.json file");
		return;
	}

	const output = window.createOutputChannel("VBA xlr8");
	const folder = folders[0];

	output.appendLine(`Working at ${folder.name}: ${folder.uri.fsPath}`);
	output.appendLine(`Working at ${folder.uri.path}`);
	

	let suscription = commands.registerCommand('vba-xlr8.compile', () => {
		// The code you place here will be executed every time your command is executed
		// Display a message box to the user
		if (proccess !== undefined) { return; }
		commands.executeCommand('setContext', vbaRunningContext, true);
		const compilerModule = context.asAbsolutePath(path.join('compiler', 'VBA.Compiler.exe'));

		proccess = child.execFile(
			compilerModule,
			['test', 'arguments'],
			{
				cwd: folder.uri.fsPath,
			},
			(error, stdout, stderr) => {
				if (error) {
					throw error;
				}
				if(stderr){
					console.log(stderr);
				}
				const editor = window.activeTextEditor;
				if (editor) {
					editor.edit((e) => {
						e.insert(editor.selection.active, stdout);
					});
				}
				output.appendLine(stdout);
			});

		proccess!.on('exit', (code) => {
			proccess = undefined;
			commands.executeCommand('setContext', vbaRunningContext, false);
		});
	});

	context.subscriptions.push(suscription);

	suscription = commands.registerCommand('vba-xlr8.pauseRun', () => {
		if (proccess) {
			proccess.kill();
			proccess = undefined;
			commands.executeCommand('setContext', vbaRunningContext, false);
		}
	});
	context.subscriptions.push(suscription);
}

// this method is called when your extension is deactivated
export function deactivate() {
	proccess?.kill();
	proccess = undefined;
 }
