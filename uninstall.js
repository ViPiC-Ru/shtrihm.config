/*! 0.1.2 удаление заданных программ или подсчёт их колличества

cscript uninstall.js [<host>] <type> <author> <name> <version> [<command>]

<host>		- имя компьютера начиная с \\ для работы с удалённым компьютером
<type>		- тип проверяемых программ x86, x64, native, computer, user или all
<author>	- регистронезависимая часть автора программы для фильтрации списка
<name>		- регистронезависимая часть названия программы для фильтрации списка
<version>	- регистронезависимая часть версии программы для фильтрации списка
<command>	- комманда которая добавляються к строке удаления программы как параметры
			  если не задавать этот параметр, то будет возврашего только колличество
			  программ ввиде кода возврата без удаления этих программ. Также можно
			  только вывести программы на экран коммандой print или csv.

 */

(function(wsh, undefined){// замыкаем что бы не сорить глобалы
	var method, param, item, items, key, keys, value, flag, count, unit, data,
		response, node, nodes, branch, branches, locator, service, registry,
		command, type = 'all', host = '', name = '', version = '',
		author = '', title = '', list = [], split = '\r\n',
		delim = ';', timeout = 1000, error = 0;
	
	// получаем параметры из коммандной строки
	if(!error){// если нету ошибок
		for(var i = 0, iLen = wsh.arguments.length; i < iLen; i++){
			value = wsh.arguments(i);// получаем очередной параметр
			switch(i){// поддерживаемые параметры коммандной строки
				case 0:// имя удалённого компьютера
					flag = !value.indexOf('\\\\');
					if(flag) host = value.substr(2);
					else type = value;
					break;
				case 1:// тип проверяемых программ
					if(flag) type = value;
					else author = value;
					break;
				case 2:// название программы
					if(flag) author = value;
					else name = value;
					break;
				case 3:// версия программы
					if(flag) name = value;
					else version = value;
					break;
				case 4:// автор программы
					if(flag) version = value;
					else command = value;
					break;
				case 5:// комманда удаления
					command = value;
					break;
			};
		};
	};
	// получаем сервис для доступа к реестру через wmi
	if(!error){// если нету ошибок
		try{// пробуем подключиться к компьютеру
			locator = new ActiveXObject('wbemScripting.Swbemlocator');
			locator.security_.impersonationLevel = 3;
			service = locator.connectServer(host, 'root\\CIMV2');
			execution = service.get('Win32_Process');// удалённый запуск процессов
			registry = locator.connectServer(host, 'root\\default').get('stdRegProv');
		}catch(e){error = 1;};
	};
	// определяем имя компьютера и разрядность системы
	if(!error){// если нету ошибок
		response = service.execQuery(
			"SELECT systemType, dnsHostName\
			 FROM Win32_ComputerSystem"
		);
		items = new Enumerator(response);
		for(count = 0; !items.atEnd(); items.moveNext()){// пробигаемся по коллекции
			item = items.item();// получаем очередной элимент коллекции
			if(item.dnsHostName) title = item.dnsHostName;
			flag = item.systemType && ~item.systemType.indexOf('64');
			count++;// увеличиваем счётчик элиментов
			// останавливаемся на первом элименте
			break;
		};
		if(count){// если есть элименты
		}else error = 2;
	};
	// формируем список веток реестра для проверки
	if(!error){// если нету ошибок
		data = {// ветки реестра
			nat: {// программ в нативном реестре
				root: 0x80000002,// HKEY_LOCAL_MACHINE
				path: 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall'
			},
			x86: {// программ в x86 реестре
				root: 0x80000002,// HKEY_LOCAL_MACHINE
				path: 'SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall'
			},
			usr: {// программ в реестре пользователя
				root: 0x80000001,// HKEY_CURRENT_USER
				path: 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall'
			}
		};
		switch(type){// поддерживаемые типы
			case 'x64':// 64 разрядные
				nodes = flag ? [data['nat']] : [];
				break;
			case 'x86':// 32 разрядные
				nodes = !flag ? [data['nat']] : [data['x86']];
				break;
			case 'native':// разрядность системы
				nodes = [data['nat']];
				break;
			case 'computer':// компьютер
				nodes = [data['nat'], data['x86']];
				break;
			case 'user':// пользователь
				nodes = [data['usr']];
				break;
			case 'all':// все
				nodes = [data['nat'], data['x86'], data['usr']];
				break;
			default:// прочии значения
				error = 3;
		};
	};
	// формируем список программ
	if(!error){// если нету ошибок
		for(var i = 0, iLen = nodes.length; i < iLen; i++){// пробигаемся по списку
			node = nodes[i];// получаем очередной элимент
			// получаем список дочерних веток
			method = registry.methods_.item('EnumKey');
			param = method.inParameters.spawnInstance_();
			param.hDefKey = node.root;
			param.sSubKeyName = node.path;
			try{// пробуем получить список веток реестра
				item = registry.execMethod_(method.name, param); 
			}catch(e){continue;};
			if(!item.returnValue){// если удалось получить список значений
				try{// пробуем получить список имён веток реестра
					branches = item.sNames.toArray();
				}catch(e){continue;};
				for(var j = 0, jLen = branches.length; j < jLen; j++){// пробигаемся по списку
					branch = branches[j];// получаем очередной элимент из списка
					// проверяем значение ключей на соответствование
					flag = true;// элимент соответствует критериям поиска
					unit = {title: title, author: author, name: name, version: version, uninstall: ''};
					data = {name: 'DisplayName', author: 'Publisher', version: 'DisplayVersion', uninstall: 'UninstallString'};
					for(var key in data){// пробигаемся по проверяемым ключам
						if(flag){// если необходимо проверить значение
							value = unit[key].toLowerCase();// преобразовываем значение
							method = registry.methods_.item('GetStringValue');
							param = method.inParameters.spawnInstance_();
							param.hDefKey = node.root;
							param.sSubKeyName = node.path + '\\' + branch;
							param.sValueName = data[key];
							item = registry.execMethod_(method.name, param); 
							if(!item.returnValue && item.sValue){// если удалось получить значение
								if(value && !~item.sValue.toLowerCase().indexOf(value)) flag = false;
								unit[key] = item.sValue;// сохраняем значение
							}else if(value) flag = false;
						}else break;
					};
					// корректируем команду на удаление
					if(flag && unit.uninstall){// если нужно выполнить
						// проверяем наличее файла
						value = unit.uninstall;
						value = value.replace(/\\/g, '\\\\');
						value = value.replace(/'/g, '"');
						response = service.execQuery(
							"SELECT name\
							 FROM CIM_DataFile\
							 WHERE name = '" + value + "'"
						);
						items = new Enumerator(response);
						for(count = 0; !items.atEnd(); items.moveNext()){// пробигаемся по коллекции
							item = items.item();// получаем очередной элимент коллекции
							unit.uninstall = '"' + unit.uninstall + '"'; 
							count++;// увеличиваем счётчик элиментов
							// останавливаемся на первом элименте
							break;
						};
						// исправляем строку удаления
						if(!count){// если это комманда на удаление
							value = unit.uninstall;
							value = value.replace('/I{', '/X{');// поправка для msi
							unit.uninstall = value;
						};
					};
					// добавляем программу в список
					if(flag) list.push(unit);
				};
			};
		};
	};
	// выполняем действие над списком программ
	if(!error){// если нету ошибок
		switch(command){// поддерживаемые команды
			case 'print':// вывести простой список программ
				items = [];// массив строчек вывода
				for(var i = 0, iLen = list.length; i < iLen; i++){// пробигаемся по списку
					unit = list[i];// получаем очередной элимент из списка
					if(unit.name){// если есть название программы
						value = unit.version ? unit.name + ' ' + unit.version : unit.name;
						items.push(value);
					};
				};
				value = items.join(split);
				if(value) wsh.echo(value);
				break;
			case 'csv':// вывести форматированный список программ
				items = [];// массив строчек вывода
				for(var i = 0, iLen = list.length; i < iLen; i++){// пробигаемся по списку
					unit = list[i];// получаем очередной элимент из списка
					nodes = [];// массив элиментов строки вывода
					for(var key in unit){// пробигаемся по ключам
						value = unit[key];// получаем очередной элимент из списка
						value = value.split(delim).join(' ');
						nodes.push(value);
					};
					value = nodes.join(delim);
					items.push(value);
				};
				value = items.join(split);
				if(value) wsh.echo(value);
				break;
			case undefined:// не выполнять удаление
				break;
			default:// выполнять удаление
				for(var i = 0, iLen = list.length; i < iLen; i++){// пробигаемся по списку
					unit = list[i];// получаем очередной элимент из списка
					value = command ? unit.uninstall + ' ' + command : unit.uninstall;
					method = execution.methods_.item('Create');
					param = method.inParameters.spawnInstance_();
					param.CommandLine = value;
					item = execution.execMethod_(method.name, param); 
					if(item.processId){// если удалось запустить выполнение
						value = item.processId;// сохраняем идентификатор
						do{// ожидаем завершение запущенного процесса
							response = service.execQuery(
								"SELECT handle\
								 FROM Win32_Process\
								 WHERE handle = '" + value + "'\
								 OR parentProcessId = '" + value + "'"
							);
							items = new Enumerator(response);
							for(count = 0; !items.atEnd(); items.moveNext()){// пробигаемся по коллекции
								item = items.item();// получаем очередной элимент коллекции
								if(item.handle && value != item.handle) value = item.handle;
								wsh.sleep(timeout);// деламем паузу между проверками
								count++;// увеличиваем счётчик элиментов
								// останавливаемся на первом элименте
								break;
							};
						}while(count);
					};
				};
				break;
		};
	};
	// завершаем сценарий кодом
	wsh.quit(list.length);
})(WSH);