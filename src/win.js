const https = require("https")
const fs = require("fs")
const Powershell = require("node-powershell")
const extract = require('extract-zip')
const { exec } = require("child_process")

exec("sc query photonservice", ( err, stdout ) => {
	let res = stdout.replace(/ +/g, '').split("\n")
	if (/STATE/.test(res[3])) {
		let state = res[3].split(":")[1].slice(1).toLowerCase()
		res = state.trim()
	} else {
		res = "error"
	}

	switch (res) {
		case "running": 
		case "start_pending": {
			runAsAdmin(`sc.exe stop photonservice; sc.exe delete photonservice`)
			break
		}
		case "stopped":
		case "stop_pending": {
			runAsAdmin(`sc.exe delete photonservice`)
			break
		}
	}
})

getFolder().then(async ( folder = `${process.env.LOCALAPPDATA}/Programs` ) => {
	process.stdout.write("\x1B[?25l")

	mkdirSafe(`${folder}/Photon`)

	fs.writeFileSync(`${folder}/Photon/ui.zip`, await getRelease("https://github.com/m4rch-xyz/photon/releases/download/v0.1.0/v0.1.0_win-x64.zip"))

	mkdirSafe(`${folder}/Photon/ui`)
	await extract(`${folder}/Photon/ui.zip`, { dir: `${folder}/Photon/ui` })
	fs.unlinkSync(`${folder}/Photon/ui.zip`)

	mkdirSafe(`${folder}/Photon/service`)
	fs.writeFileSync(`${folder}/Photon/service/service.exe`, await getRelease("https://github.com/m4rch-xyz/photon-api/releases/download/v0.1.0/v0.1.0_win-x64.exe"))
	fs.writeFileSync(`${folder}/Photon/service/winsw.exe`, await getRelease("https://github.com/winsw/winsw/releases/download/v2.11.0/WinSW-x64.exe"))

	fs.writeFileSync(`${folder}/Photon/service/winsw.xml`, [
		"<service>",
			"\t<id>photonservice</id>",
			"\t<name>PhotonService</name>",
			"\t<description>Backgroundworker for Photon Application.</description>",
			`\t<executable>${folder}\\Photon\\service\\service.exe</executable>`,
			"\t<logmode>rotate</logmode>",
			"\t<stoptimeout>30sec</stoptimeout>",
		"</service>"
	].join("\n"))

	runAsAdmin(`${folder}\\Photon\\service\\winsw.exe install; ${folder}\\Photon\\service\\winsw.exe start`)

	mkdirSafe(`${process.env.ProgramData}/Photon`, false)
	mkdirSafe(`${process.env.ProgramData}/Photon/data`, false)
	mkdirSafe(`${process.env.ProgramData}/Photon/profiles`, false)

	mkfileSafe(`${process.env.ProgramData}/Photon/data/last.json`, JSON.stringify({
		"device": [],
		"syncZones": {}
	}))

	mkfileSafe(`${process.env.ProgramData}/Photon/data/settings.json`, JSON.stringify({
		"startupProfile": false,
		"shutdownProfile": false
	}))

 	generateLink(folder)

	console.log("installation finished")
})

function getFolder () {
	return new Promise(( resolve, reject ) => {
		switch (process.platform) {
			case "win32": {
				const ps = new Powershell({
					executionPolicy: 'Bypass',
					noProfile: true
				});
				ps.addCommand(`
					[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
	
					$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
					$foldername.Description = "Select a folder"
					$foldername.SelectedPath = "${process.env.LOCALAPPDATA}\\Programs\\"
	
					if($foldername.ShowDialog() -eq "OK")
					{
						$folder += $foldername.SelectedPath
					}
					echo $folder
				`)
				ps.invoke()
					.then(( output ) => { ps.dispose(), resolve(output.trim()) })
					.catch(( err ) => { ps.dispose(), reject(err) })
				break
			}
			case "linux": {

			}
		}
	}) 
}

function generateLink ( folder ) {
	const ps = new Powershell({
		executionPolicy: 'Bypass',
		noProfile: true
	})

	ps.addCommand(`
		$objShell = New-Object -ComObject ("WScript.Shell")
		$objShortCut = $objShell.CreateShortcut("${process.env.APPDATA}\\Microsoft\\Windows\\Start Menu\\Programs\\Photon.lnk")
		$objShortCut.TargetPath = "${folder}\\Photon\\ui\\photon-win32-x64\\photon.exe"
		$objShortCut.Save()
	`)

	ps.invoke()
		.then(() => ps.dispose())
		.catch(() => ps.dispose())
}

function getRelease ( link ) {
	return new Promise(( resolve ) => {
		https.get(link, ( res ) => {
			res.on("data", () => {})

			res.on("end", () => {
				https.get(res.headers.location, ( res ) => {
					let bar = loading(res.headers["content-length"])

					let arr = []
					res.on("data", ( data ) => {
						arr.push(data)
						bar.add(data.length).log()
					})

					res.on("end", () => {
						resolve(Buffer.concat(arr))
						process.stdout.write("\n")
					})
				})
			})
		})
	})
}

function mkdirSafe ( path, overwrite = true ) {
	if (fs.existsSync(path) && !fs.lstatSync(path).isDirectory() || !fs.existsSync(path)) {
		fs.mkdirSync(path)
	} else if (overwrite && fs.existsSync(path) && fs.lstatSync(path).isDirectory()) {
		fs.rmdirSync(path, { recursive: true })
		fs.mkdirSync(path)
	}
}

function mkfileSafe ( path, value, overwrite = true ) {
	if (overwrite && fs.existsSync(path) && !fs.lstatSync(path).isFile() || !fs.existsSync(path)) {
		fs.writeFileSync(path, value)
	}
}

async function runAsAdmin ( command ) {
	const shell = new Powershell({})

	await shell.addCommand('Start-Process')
	await shell.addArgument('PowerShell')

	await shell.addArgument('-Verb')
	await shell.addArgument('RunAs')

	await shell.addArgument('-WindowStyle')
	await shell.addArgument('Hidden')

	await shell.addArgument('-PassThru')

	await shell.addArgument('-Wait')

	await shell.addArgument('-ArgumentList')
	await shell.addArgument(`\"${command}\"`)

	await shell.invoke()
	return await shell.dispose()
}

function loading ( num ) {
	const { stdout } = process

	let current = 0
	let end = +num

	return Object.freeze({
		add ( amount ) {
			current += amount
			return this
		},
		log () {
			stdout.cursorTo(0)
			stdout.write(`downloading content: \x1b[33m${(current / end * 100).toFixed(2)}%\x1b[39m \x1b[32m(${(current / 1_000_000).toFixed(2)}MB)\x1b[39m`)
		}
	})
}

process.on("exit", () => process.stdout.write("\x1B[?25h"))
