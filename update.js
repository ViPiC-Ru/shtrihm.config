/*! 0.2.1 обновление касс в связи со сменой ндс

cscript update.min.js [[[<install>] <config>] <license>]

<install>	- относительный путь к файлу установки драйвера или false для пропуска.
<config>	- true если нужно внести изменения в настройки кассы или false для пропуска.
<license>	- относительный путь к файлу лицензий для касс или false для пропуска.

 */

(function(wsh, undefined){// замыкаем что бы не сорить глобалы
	var list, value, command, shell, fso, ts, driver, item, node, nodes, flag,
		install, license, config = null, isConnect = null, isSilent = true,
		dLine = '\r\n', dValue = '\t', error = 0;
	
	shell = new ActiveXObject('WScript.Shell');
	fso = new ActiveXObject('Scripting.FileSystemObject'); 
	// получаем путь к файлу установки драйвера
	if(!error){// если нету ошибок
		if(0 < wsh.arguments.length){// если передан параметр
			value = wsh.arguments(0);
			if(value && 'false' != value.toLowerCase()){// если задано
				install = fso.getAbsolutePathName(value);
			};
		};
	};
	// получаем флаг для изменения настройки кассы
	if(!error){// если нету ошибок
		if(1 < wsh.arguments.length){// если передан параметр
			value = wsh.arguments(1);
			config = 'true' == value.toLowerCase();
		};
	};
	// получаем путь к файлу лицензий для касс
	if(!error){// если нету ошибок
		if(2 < wsh.arguments.length){// если передан параметр
			value = wsh.arguments(2);
			if(value && 'false' != value.toLowerCase()){// если задано
				license = fso.getAbsolutePathName(value);
			};
		};
	};
	// проверяем наличее файла установки драйвера
	if(!error && install){// если нужно выполнить
		if(fso.fileExists(install)){// если файл существует
		}else error = 1;
	};
	// проверяем наличее файла лицензий для касс
	if(!error && license){// если нужно выполнить
		if(fso.fileExists(license)){// если файл существует
		}else error = 2;
	};
	// вычисляем вспомогательные переменные
	if(!error){// если нету ошибок
		isSilent = !install && !config && !license;
	};
	// показываем начальное сообщение пользователю
	if(!isSilent){// если нужно выполнить
		value = // сообщение для пользователя
			'Через минуту на компьютер будет установлено обновление для ККМ. ' +
			'Нужно будет закрыть кассу программы еФарма. Закрывать смену при этом не нужно. ' +
			'Установка займёт три минуты. После этого вы сможете работать.';
		command = 'shutdown /r /t 60 /c "' + value  + '"';
		shell.run(command, 0, false);
		wsh.sleep(30 * 1000);
		command = 'shutdown /a';
		shell.run(command, 0, true);
	};	
	// принудительно завершаем работу кассовой програмы
	if(!error && !isSilent){// если нужно выполнить
		command = 'taskkill /F /IM ePlus.ARMCasherNew.exe /T';
		shell.run(command, 0, true);
		wsh.sleep(2 * 1000);
	};
	// выполняем удаление и установку драйвера
	if(install){// если нужно обновить драйвер
		// удаляем все установленные версии драйвера
		if(!error){// если нету ошибок
			value = 'all "Штрих" "Драйвер ФР" "" "/verysilent"';
			command = 'cscript uninstall.js ' + value;
			shell.run(command, 0, true);
		};
		// удаляем возможные старые версии
		if(!error){// если нету ошибок
			list = [// список корневых папок для поиска дочерних
				{path: 'C:\\Program Files\\SHTRIH-M', filter: 'DrvFR'},
				{path: 'C:\\Program Files (x86)\\SHTRIH-M', filter: 'DrvFR'},
				{path: 'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\ШТРИХ-М', filter: 'ФР'}
			];
			for(var i = 0, iLen = list.length; i < iLen && !error; i++){
				item = list[i];// получаем очередной элимент
				if(fso.folderExists(item.path)){// если папка существует
					node = fso.getFolder(item.path);
					nodes = new Enumerator(node.subFolders);
					while(!nodes.atEnd()){// пока не достигнут конец списка
						node = nodes.item();// получаем очередной элимент
						if(~node.name.indexOf(item.filter)){// содержит строку фильтра
							try{// пробуем удалить полученный элимент
								node.Delete(true);
							}catch(e){};
						};
						nodes.moveNext();
					};
				};
			};
		};
		// устанавливаем последнюю версию драйвера
		if(!error){// если нету ошибок
			command = '"' + install + '" /verysilent';
			value = shell.run(command, 0, true);
			if(!value){// если комманда выполнена успешно
			}else error = 3;
		};
	};
	// готовимся к взаимодействию с кассой
	if(config || license){// если нужно взаимодействать с кассой
		// создаём объект для взаимодейсивия с кассой
		if(!error){// если нету ошибок
			try{// пробуем подключиться к кассе
				driver = new ActiveXObject('Addin.DrvFR');
			}catch(e){error = 4;};
		};
		// подключаемся к кассе
		if(!error){// если нету ошибок
			driver.Password = 30;
			driver.GetECRStatus();
			isConnect = false;
			switch(driver.ResultCode){
				case  0: isConnect = true; break;	// ккм доступна
				case -1: break;						// ккм не подключена
				case -3: error = 5; break;			// ккм занята
				default: error = 6;					// другие ошибки
			};
		};
	};
	// изменяем данные в таблице кассы
	if(!error && config && isConnect){// если нужно выполнить
		driver.Password = 30;
		list = [// изменяемые поля таблицы
			// тип и режим кассы
			{table:  1, row: 1, field:  7, value: 2},// отрезка чека
			// сетевой адрес
			{table: 16, row: 1, field:  1, value: 0},// static ip
			// региональные настройки
			{table: 17, row: 1, field:  3, value: 2},// режим исчисления скидок
			{table: 17, row: 1, field: 10, value: 1},// печать параметров офд в чеках
			{table: 17, row: 1, field: 12, value: 7},// печать реквизитов пользователя
			{table: 17, row: 1, field: 17, value: 2},// формат фд
			// fiscal storage
			{table: 18, row: 1, field:  7, value: 'ГБУ МО "Мособлмедсервис"'},// user
			// удаленный мониторинг и администрирование
			{table: 23, row: 1, field:  1, value: 1},// работать с сервером ско
			{table: 23, row: 1, field:  5, value: 1},// разрешить автообновление
			{table: 23, row: 1, field:  6, value: 1} // однократное обновление
		];
		for(var i = 0, iLen = list.length; i < iLen && !error; i++){
			item =  list[i];// получаем очередной элимен
			// читаем текущее значение в таблице
			driver.TableNumber = item.table;
			driver.RowNumber = item.row;
			driver.FieldNumber = item.field;
			driver.GetFieldStruct();
			if(!driver.ResultCode){// если данные получены
				driver.ReadTable();// получаем данные
				if(!driver.ResultCode){// если данные получены
					if(driver.FieldType) value = driver.ValueOfFieldString;
					else value = driver.ValueOfFieldInteger;
					if(value != item.value){// если нужно изменить данные
						// изменяем значение в таблицы
						if(driver.FieldType) driver.ValueOfFieldString = item.value;
						else driver.ValueOfFieldInteger = item.value;
						driver.WriteTable();// изменяем данные
						if(!driver.ResultCode){// если данные изменены
						}else error = 9;
					};
				}else error = 8;
			}else error = 7;
		};
	};
	// выполняем дейсвия по активации лицензии
	if(license){// если нужно активировать лицензию
		// получаем содержимое файла лицензий
		if(!error){// если нету ошибок
			ts = fso.openTextFile(license, 1);
			if(!ts.atEndOfStream){// если файл не пуст
				value = ts.readAll();
			}else error = 10;
			ts.close();
		};
		// преобразовываем содержимое в список лицензий
		if(!error){// если нету ошибок
			list = value.split(dLine);
			for(var i = 0, iLen = list.length; i < iLen && !error; i++){
				list[i] = list[i].split(dValue);
				if(3 == list[i].length){// лицензии успешно разделены
					item = {// элимент данных
						serial: list[i][0],		// серийный номер
						license: list[i][1],	// лицензия
						signature: list[i][2]	// подпись
					};
					list[i] = item;
				};				
			};
		};
		// ищем и активируем лицензию
		if(!error){// если нету ошибок
			driver.Password = 30;
			// получаем длинный заводской номер
			driver.ReadSerialNumber();
			if(!driver.ResultCode){// если данные получены
				flag = false;// активирована ли лицензия
				for(var i = 0, iLen = list.length; i < iLen && !error; i++){
					item =  list[i];// получаем очередной элимен
					if(driver.SerialNumber == item.serial){// если найдена лицензия
						driver.License = item.license;
						driver.DigitalSign = item.signature;
						// активируем лицензию на кассе
						driver.WriteFeatureLicenses();
						if(!driver.ResultCode){// если лицензия активирована
							flag = true;// лицензия активирована
						}else error = 12;
					};
				};
			}else error = 11;
		};
		// проверяем активирована ли лицензия
		if(!error){// если нету ошибок
			if(flag){// если лицензия активирована
			}else error = 13;
		};
	};
	// показываем конечное сообщение пользователю
	if(!isSilent){// если нужно выполнить
		value = // основное сообщение для пользователя
			'Установка обновления завершена. Можете продолжать работу.';
		if(config) value += ' ' + // дополнительное сообщение для пользователя
			'Вечером после закрытия смены, закройте кассу программы еФарма, ' +
			'но не выключайте сам фискальный аппарат, на него будут установлены обновления. ' + 
			'На следующий рабочий день, после открытия смены, после первой продажи ' +
			'необходимо сделать внесение в кассу программы еФарма, т.к. после установки обновений ' +
			'это значение обнулится. При возникновении трудностей создайте заявку в ИТ отдел.';
		command = 'shutdown /r /t 60 /c "' + value  + '"';
		shell.run(command, 0, false);
		wsh.sleep(45 * 1000);
		command = 'shutdown /a';
		shell.run(command, 0, true);
	};	
	// завершаем сценарий кодом
	wsh.quit(error);
})(WSH);