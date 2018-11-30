/*! 0.1.1 обновление касс в связи со сменой ндс */

(function(wsh, undefined){// замыкаем что бы не сорить глобалы
	var list, value, command, shell, driver, item, flag, error = 0;
	
	shell = new ActiveXObject('WScript.Shell');
	// показываем сообщение пользователю
	value = // сообщение для пользователя
		'Через минуту на компьютер будет установлено обновление для ККМ. ' +
		'Нужно будет закрыть кассу программы еФарма. Закрывать смену при этом не нужно. ' +
		'Установка займёт три минуты. После этого вы сможете работать.';
	command = 'shutdown /r /t 60 /c "' + value  + '"';
	shell.run(command, 0, false);
	wsh.sleep(30 * 1000);
	command = 'shutdown /a';
	shell.run(command, 0, true);
	// принудительно завершаем работу кассовой програмы
	if(!error){// если нету ошибок
		command = 'taskkill /F /IM ePlus.ARMCasherNew.exe /T';
		shell.run(command, 0, true);
	};
	// удаляем все установленные версии драйвера
	if(!error){// если нету ошибок
		value = 'all "Штрих" "Драйвер ФР" "" "/verysilent"';
		command = 'cscript uninstall.js ' + value;
		shell.run(command, 0, true);
	};
	// устанавливаем последнюю версию драйвера
	if(!error){// если нету ошибок
		command = 'driver.exe /verysilent';
		value = shell.run(command, 0, true);
		if(!value){// если комманда выполнена успешно
		}else error = 1;
	};
	// создаём объект для зваимодейсивия с кассой
	if(!error){// если нету ошибок
		try{// пробуем подключиться к кассе
			driver = new ActiveXObject('Addin.DrvFR');
		}catch(e){error = 2;};
	};
	// подключаемся к кассе
	if(!error){// если нету ошибок
		driver.Password = 30;
		driver.GetECRStatus();
		switch(driver.ResultCode){
			case  0: flag = true; break;	// ккм доступна
			case -1: flag = false; break;	// ккм не подключена
			case -3: error = 3; break;		// ккм занята
			default: error = 4;				// другие ошибки
		};
	};
	// изменяем данные в таблице кассы
	if(!error && flag){// если нужно выполнить
		driver.Password = 30;
		list = [// изменяемые поля таблицы
			// региональные настройки
			{table: 17, row: 1, field:  3, value: 2},// режим исчисления скидок
			{table: 17, row: 1, field: 10, value: 1},// печать параметров офд в чеках
			{table: 17, row: 1, field: 12, value: 7},// печать реквизитов пользователя
			{table: 17, row: 1, field: 17, value: 2},// формат фд
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
						}else error = 7;
					};
				}else error = 6;
			}else error = 5;
		};
	};
	// показываем сообщение пользователю
	value = // сообщение для пользователя
		'Установка обновления завершена. Можете продолжать работу. Вечером после закрытия смены, ' +
		'закройте кассу программы еФарма, но не выключайте сам фискальный аппарат, на него будут установлены обновления. ' +
		'На следующий рабочий день до открытия смены позвоните в ИТ отдел.';
	command = 'shutdown /r /t 60 /c "' + value  + '"';
	shell.run(command, 0, false);
	wsh.sleep(30 * 1000);
	command = 'shutdown /a';
	shell.run(command, 0, true);
	// завершаем сценарий кодом
	wsh.quit(error);
})(WSH);