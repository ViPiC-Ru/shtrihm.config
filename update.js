/*! 0.2.6 обновление касс и изменение настроек

cscript update.min.js <install> <license> <variable> <logotype> <table> <silent>

<install>	- относительный путь к файлу установки драйвера или false для пропуска.
<license>	- относительный путь к файлу лицензий для касс или false для пропуска.
<variable>	- относительный путь к файлу с переменными или false для пропуска.
<logotype>	- относительный путь к BMP файлу логотипа для печати вверху чека.
<table>		- относительный путь к файлу таблиц для импорта в формате Штрих-М.
<silent>	- выполнять тихую установку без остановки работы пользователя

 */

(function (wsh, undefined) {// замыкаем что бы не сорить глобалы
	var value, key, list, data, name, char, command, shell, fso, stream, network, driver, item,
		node, nodes, flag, image, install, license, variable, logotype, table, template,
		dec2hex, silent = false, dLine = '\r\n', dValue = '\t', dCell = ',', error = 0;

	/**
	 * Заполняет шаблон данными из объекта.
	 * @param {string} pattern - Шаблона для подстановки данных.
	 * @param {string} data - Объект с данными для подстановки.
	 * @returns {string} Заполненный шаблон с данными.
	 */

	template = function (pattern, data) {
		var list, fragments = [], delim = '|', select = '%';

		if (pattern && data) {// если указаны обязательные параметры
			fragments = pattern.split(delim);
			for (var iLen = fragments.length, i = iLen - 1; -1 < i; i--) {
				if (fragments[i]) {// если фрагмент не пустой
					list = fragments[i].split(select);
					if (list.length % 2) {// если в шаблоне переменные
						for (var j = 1, jLen = list.length, flag = true; j < jLen; j += 2) {
							if (list[j] in data) {// если задано значение
								list[j] = data[list[j]];
							} else flag = false;
						};
						if (flag) fragments[i] = list.join('');
						else fragments.splice(i, 1);
					} else fragments.splice(i, 1);
				} else if (i && i < iLen - 1) fragments[i] = delim;
			};
		};
		return fragments.join('');
	};

	/**
	 * Переводит число из десятичного в шестнадцатеричное.
	 * @param {number} dec - Число в десячичной записе.
	 * @param {number} [length] - Длина возврашаемого значения.
	 * @returns {string} Число в шестнадцатеричной записе.
	 */

	dec2hex = function (dec, length) {
		var value, hex = '', chars = '0123456789ABCDEF';

		value = Number(dec);
		// переводим число в другую систему
		do {// циклически делим число на основание
			hex = chars.charAt(value % chars.length) + hex;
			value = Math.floor(value / chars.length);
		} while (value);
		// добовляем ведущие нули
		while (hex.length < length) hex = chars.charAt(0) + hex;
		//возвращаем результат
		return hex;
	};

	shell = new ActiveXObject('WScript.Shell');
	fso = new ActiveXObject('Scripting.FileSystemObject');
	// получаем путь к файлу установки драйвера
	if (!error) {// если нету ошибок
		if (0 < wsh.arguments.length) {// если передан параметр
			value = wsh.arguments(0);
			if (value && 'false' != value.toLowerCase()) {// если задано
				install = fso.getAbsolutePathName(value);
			};
		};
	};
	// получаем путь к файлу лицензий для касс
	if (!error) {// если нету ошибок
		if (1 < wsh.arguments.length) {// если передан параметр
			value = wsh.arguments(1);
			if (value && 'false' != value.toLowerCase()) {// если задано
				license = fso.getAbsolutePathName(value);
			};
		};
	};
	// получаем путь к файлу переменных
	if (!error) {// если нету ошибок
		if (2 < wsh.arguments.length) {// если передан параметр
			value = wsh.arguments(2);
			if (value && 'false' != value.toLowerCase()) {// если задано
				variable = fso.getAbsolutePathName(value);
			};
		};
	};
	// получаем путь к файлу логотипа
	if (!error) {// если нету ошибок
		if (3 < wsh.arguments.length) {// если передан параметр
			value = wsh.arguments(3);
			if (value && 'false' != value.toLowerCase()) {// если задано
				logotype = fso.getAbsolutePathName(value);
			};
		};
	};
	// получаем путь к файлу таблиц для импорта
	if (!error) {// если нету ошибок
		if (4 < wsh.arguments.length) {// если передан параметр
			value = wsh.arguments(4);
			if (value && 'false' != value.toLowerCase()) {// если задано
				table = fso.getAbsolutePathName(value);
			};
		};
	};
	// получаем параметр тихой установки
	if (!error) {// если нету ошибок
		if (5 < wsh.arguments.length) {// если передан параметр
			value = wsh.arguments(5);
			silent = 'true' == value.toLowerCase();
		};
	};
	// проверяем наличее файла установки драйвера
	if (!error && install) {// если нужно выполнить
		if (fso.fileExists(install)) {// если файл существует
		} else error = 1;
	};
	// проверяем наличее файла лицензий для касс
	if (!error && license) {// если нужно выполнить
		if (fso.fileExists(license)) {// если файл существует
		} else error = 2;
	};
	// проверяем наличее файла с переменными
	if (!error && variable) {// если нужно выполнить
		if (fso.fileExists(variable)) {// если файл существует
		} else error = 3;
	};
	// проверяем наличее файла логотипа
	if (!error && logotype) {// если нужно выполнить
		if (fso.fileExists(logotype)) {// если файл существует
		} else error = 4;
	};
	// проверяем наличее файла таблиц для импорта
	if (!error && table) {// если нужно выполнить
		if (fso.fileExists(table)) {// если файл существует
		} else error = 5;
	};
	// прерываем работу пользователя
	if (!silent) {// если можно прервать работу пользователя
		// показываем начальное сообщение пользователю
		if (!error) {// если нету ошибок
			value = // сообщение для пользователя
				'Через 2 минуты на компьютер будет установлено обновление для ККМ. ' +
				'Нужно будет закрыть кассу программы еФарма. Закрывать смену при этом не нужно. ' +
				'Установка займёт 3 минуты. После этого вы сможете работать.';
			command = 'shutdown /r /t 60 /c "' + value + '"';
			shell.run(command, 0, false);
			wsh.sleep(30 * 1000);
			command = 'shutdown /a';
			shell.run(command, 0, true);
			wsh.sleep(90 * 1000);
		};
		// принудительно завершаем работу кассовой програмы
		if (!error) {// если нету ошибок
			command = 'taskkill /F /IM ePlus.ARMCasherNew.exe /T';
			shell.run(command, 0, true);
			wsh.sleep(2 * 1000);
		};
	};
	// выполняем удаление и установку драйвера
	if (install) {// если нужно обновить драйвер
		// удаляем все установленные версии драйвера
		if (!error) {// если нету ошибок
			value = 'all "Штрих" "Драйвер ФР" "" "/verysilent"';
			command = 'cscript uninstall.js ' + value;
			shell.run(command, 0, true);
		};
		// удаляем возможные старые версии
		if (!error) {// если нету ошибок
			list = [// список корневых папок для поиска дочерних
				{ path: 'C:\\Program Files\\SHTRIH-M', filter: 'DrvFR' },
				{ path: 'C:\\Program Files (x86)\\SHTRIH-M', filter: 'DrvFR' },
				{ path: 'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\ШТРИХ-М', filter: 'ФР' }
			];
			for (var i = 0, iLen = list.length; i < iLen; i++) {
				item = list[i];// получаем очередной элимент
				if (fso.folderExists(item.path)) {// если папка существует
					node = fso.getFolder(item.path);
					nodes = new Enumerator(node.subFolders);
					while (!nodes.atEnd()) {// пока не достигнут конец списка
						node = nodes.item();// получаем очередной элимент
						if (~node.name.indexOf(item.filter)) {// содержит строку фильтра
							try {// пробуем удалить полученный элимент
								node.Delete(true);
							} catch (e) { };
						};
						nodes.moveNext();
					};
				};
			};
		};
		// устанавливаем последнюю версию драйвера
		if (!error) {// если нету ошибок
			command = '"' + install + '" /verysilent';
			value = shell.run(command, 0, true);
			if (!value) {// если комманда выполнена успешно
			} else error = 6;
		};
	};
	// готовимся к взаимодействию с кассой
	if (logotype || table || license) {// если нужно взаимодействать с кассой
		// создаём объект для взаимодейсивия с кассой
		if (!error) {// если нету ошибок
			try {// пробуем подключиться к кассе
				driver = new ActiveXObject('Addin.DrvFR');
			} catch (e) { error = 7; };
		};
		// подключаемся к кассе
		if (!error) {// если нету ошибок
			driver.Password = 30;
			driver.GetECRStatus();
			switch (driver.ResultCode) {
				case 0: break;				// ккм доступна
				case -1: error = 8; break;	// ккм не подключена
				case -3: error = 9; break;	// ккм занята
				default: error = 10;		// другие ошибки
			};
		};
	};
	// выполняем активацию лицензии
	if (license) {// если нужно активировать лицензию
		// получаем содержимое файла
		if (!error) {// если нету ошибок
			stream = fso.openTextFile(license, 1);
			if (!stream.atEndOfStream) {// если файл не пуст
				data = stream.readAll();
			} else error = 11;
			stream.close();
		};
		// преобразовываем содержимое в список
		if (!error) {// если нету ошибок
			list = data.split(dLine);
			for (var i = 0, iLen = list.length; i < iLen; i++) {
				list[i] = list[i].split(dValue);// разделяем значения в строке
				if (3 == list[i].length) {// лицензии успешно разделены
					item = {// элимент данных
						serial: list[i][0],		// серийный номер
						license: list[i][1],	// лицензия
						signature: list[i][2]	// подпись
					};
					list[i] = item;
				} else list[i] = null;
			};
		};
		// ищем и активируем лицензию
		if (!error) {// если нету ошибок
			// получаем длинный заводской номер
			driver.ReadSerialNumber();
			if (!driver.ResultCode) {// если данные получены
				flag = false;// активирована ли лицензия
				for (var i = 0, iLen = list.length; i < iLen && !error; i++) {
					item = list[i];// получаем очередной элимен
					if (item && driver.SerialNumber == item.serial) {// если найдена лицензия
						driver.License = item.license;
						driver.DigitalSign = item.signature;
						// активируем лицензию на кассе
						driver.WriteFeatureLicenses();
						if (!driver.ResultCode) {// если лицензия активирована
							flag = true;// лицензия активирована
						} else error = 13;
					};
				};
			} else error = 12;
		};
		// проверяем активирована ли лицензия
		if (!error) {// если нету ошибок
			if (flag) {// если лицензия активирована
			} else error = 14;
		};
	};
	// получаем значение переменных из файла
	if (variable) {// если нужно загрузить переменные
		// получаем имя копьютера для поиска переменных
		if (!error) {// если нету ошибок
			network = new ActiveXObject('WScript.Network');
			name = network.computerName.toLowerCase();
		};
		// получаем содержимое файла
		if (!error) {// если нету ошибок
			stream = fso.openTextFile(variable, 1);
			if (!stream.atEndOfStream) {// если файл не пуст
				data = stream.readAll();
			} else error = 15;
			stream.close();
			variable = null;
		};
		// преобразовываем содержимое в список
		if (!error) {// если нету ошибок
			list = data.split(dLine);
			flag = false;// найдены ли переменные для этого компьютера
			for (var i = 0, iLen = list.length; i < iLen && !flag; i++) {
				list[i] = list[i].split(dValue);// разделяем значения в строке
				if (i) {// если это строка с данными
					item = {};// элимент данных
					// формируем элимент данных по ключам из первой строки
					for (var j = 0, jLen = list[0].length; j < jLen; j++) {
						key = list[0][j];// получаем очередной ключ
						value = list[i][j];// получаем очередное значение
						if (!j && value) value = value.toLowerCase();
						if (key && value) item[key] = value;
					};
					// запоминаем элимент данных
					if (!name.indexOf(item.name)) {// если найдены переменные
						variable = item;// добавляем переменные
						flag = true;// переменные найдены
					};
				};
			};
		};
	};
	// добавляем логотип в клише
	if (logotype) {// если нужно выполнить
		// получаем данные логотипа
		if (!error) {// если нету ошибок
			image = new ActiveXObject('WIA.ImageFile');
			image.loadFile(logotype);// читаем файл
			list = [];// список значений для байтов
			for (var y = 0, yLen = 128; y < yLen; y++) {// высота
				for (var x = 0, xLen = 512; x < xLen; x++) {// ширина
					// вычисляем значение пиксела
					if (x < image.width && y < image.height) {// если есть пиксел
						value = image.ARGBData(x + y * image.width + 1);
						value = -16777216 == value ? 1 : 0;// чёрный цвет
					} else value = 0;
					// формируем данные 8 пиксилов
					char = x % 8 ? char : 0;
					char += value ? Math.pow(2, x % 8) : 0;
					if (7 == x % 8) list.push(dec2hex(char, 2));
				};
			};
		};
		// загружаем изображение
		if (!error) {// если нету ошибок
			driver.LineNumber = 0;
			driver.LineDataHex = list.join(' ');
			driver.WideLoadLineData();
			if (!driver.ResultCode) {// если изображение загружено
			} else error = 15;
		};
	};
	// добавляем данные в таблицы
	if (table) {// если нужно выполнить
		// получаем содержимое файла
		if (!error) {// если нету ошибок
			stream = fso.openTextFile(table, 1, false, -1);
			if (!stream.atEndOfStream) {// если файл не пуст
				data = stream.readAll();
			} else error = 16;
			stream.close();
		};
		// преобразовываем содержимое в список
		if (!error) {// если нету ошибок
			list = data.split(dLine);
			for (var i = 0, iLen = list.length; i < iLen; i++) {
				value = list[i];// сохраняем значение строки
				list[i] = list[i].split(dCell);// разделяем значения в строке
				if (list[i].length > 3 && value.indexOf('//')) {
					value = value.split("','")[1] || '';
					value = value.substr(0, value.length - 1);
					value = template(value, variable || {});
					item = {// элимент данных
						table: list[i][0],	// таблица
						row: list[i][1],	// ряд
						field: list[i][2],	// поле
						value: value		// значение
					};
					list[i] = item;
				} else list[i] = null;
			};
		};
		// выполняем импорт значений таблиц
		if (!error) {// если нету ошибок
			for (var i = 0, iLen = list.length; i < iLen; i++) {
				item = list[i];// получаем очередной элимен
				// читаем текущее значение в таблице
				if (item) {// если не пустая строка таблицы
					driver.TableNumber = item.table;
					driver.RowNumber = item.row;
					driver.FieldNumber = item.field;
					driver.GetFieldStruct();
					if (!driver.ResultCode) {// если данные получены
						driver.ReadTable();// получаем данные
						if (!driver.ResultCode) {// если данные получены
							if (driver.FieldType) value = driver.ValueOfFieldString;
							else value = driver.ValueOfFieldInteger;
							if (value != item.value) {// если нужно изменить данные
								// изменяем значение в таблицы
								if (driver.FieldType) driver.ValueOfFieldString = item.value;
								else driver.ValueOfFieldInteger = item.value;
								driver.WriteTable();// изменяем данные
							};
						};
					};
				};
			};
		};
	};
	// возобнавляем работу пользователя
	if (!silent) {// если нужно возобновить работу пользователя
		// показываем конечное сообщение пользователю
		value = 'Установка обновления для ККМ завершена. Можете продолжать работу.';
		command = 'shutdown /r /t 60 /c "' + value + '"';
		shell.run(command, 0, false);
		wsh.sleep(30 * 1000);
		command = 'shutdown /a';
		shell.run(command, 0, true);
	};
	// завершаем сценарий кодом
	wsh.quit(error);
})(WSH);