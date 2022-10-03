import { commands, ExtensionContext, window } from "vscode";
import * as child from 'child_process';
import * as path from 'path';


const vbaRunningContext = 'vbaxrl8run';
let proccess: child.ChildProcess | undefined = undefined;

export function activate(context: ExtensionContext) {

	// Use the console to output diagnostic information (console.log) and errors (console.error)
	// This line of code will only be executed once when your extension is activated
	console.log('Congratulations, your extension "vba-xlr8" is now active!');

	// The command has been defined in the package.json file
	// Now provide the implementation of the command with registerCommand
	// The commandId parameter must match the command field in package.json
	let suscription = commands.registerCommand('vba-xlr8.vbaRun', () => {
		// The code you place here will be executed every time your command is executed
		// Display a message box to the user
		if (proccess !== undefined) { return; }
		commands.executeCommand('setContext', vbaRunningContext, true);
		const compilerModule = context.asAbsolutePath(path.join('compiler', 'VBA.Compiler.exe'));
		console.log(compilerModule);

		proccess = child.execFile(
			compilerModule,
			['la', 'puta'],
			(error, stdout, stderr) => {
				if (error) {
					throw error;
				}
				const editor = window.activeTextEditor;
				if (editor) {
					editor.edit((e) => {
						e.insert(editor.selection.active, stdout);
					});
				}
				console.log(stdout);
			});

		proccess!.on('exit', (code) => {
			console.log('process exit with code: ' + code);
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
export function deactivate() { }
