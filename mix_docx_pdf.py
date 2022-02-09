# mix_docx_pdf

# from glob import glob
import os
from pathlib import Path
from time import localtime, strftime


from docx2pdf import convert  # pip install docx2pdf

# pip install borb
from borb.pdf.document import Document
from borb.pdf.pdf import PDF

import yaml


ERRORS_FILE = "Ошибки.txt"
WARNINGS_FILE = "Предупреждения.txt"
CONFIG_FILE = "mix_config.yml"

WIN_LONG_PATH_PREFIX = '\\\\?\\'  # \\?\  - специальный префикс для использования длинных путей в Windows


class jsObj(dict):
	'JS-object-like dict (access to "foo": obj.foo as well as obj["foo"])'
	# __getattr__ = dict.__getitem__
	__getattr__ = dict.get   # interface modification
	__setattr__ = dict.__setitem__


def listdir(dir_path, pattern='*.*') -> list[Path]:
	return [*Path(dir_path).glob(pattern)]


def cut_dir_ext(*paths):
	'return pure file name(s), without any/dirs/ and .ext'
	if len(paths) == 1:
		p = paths[0]
		if isinstance(p, (Path)):
			return p.stem
		else:
			paths = p
	return [p.stem for p in paths]


# How to use special prefix with pathlib:  https://stackoverflow.com/questions/55815617/pathlib-path-rglob-fails-on-long-file-paths-in-windows
def longpath(path):
    normalized = os.fspath(Path(path).resolve())
    if not normalized.startswith(WIN_LONG_PATH_PREFIX):
        normalized = WIN_LONG_PATH_PREFIX + normalized
    return Path(normalized)


def strip_path_prefix(p:Path or str) -> str:
	p = str(p)
	if p.startswith(WIN_LONG_PATH_PREFIX):
		p = p[len(WIN_LONG_PATH_PREFIX):]
	return p


def create_dir(p:Path or str) -> Path:
	p = Path(longpath(p))
	p.mkdir(parents=True, exist_ok=True)
	return p



def _log(*args):
	'_log(text, [iteration, total, ] progress_callback)'
	progress_callback = args[-1]
	args = args[:-1]
	if args and progress_callback:
		# text [, iteration, total]
		progress_callback(*args)



def find_files(config: dict={}):
	config = jsObj(config)
	log = lambda *a: _log(*a, config.progress_callback)

	docx_src_files = config.docx_src_files or ()
	pdf_src_files  = config.pdf_src_files  or ()

	docx_src_dir = config.docx_src_dir  # or None
	pdf_src_dir  = config.pdf_src_dir   # or None
	log_dir  = create_dir(config.log_dir)


	if not docx_src_files and docx_src_dir:
		log("Читаю файлы ...")
		docx_src_files = listdir(longpath(docx_src_dir), '*.docx')
		log("%d файлов DOCX." % len(docx_src_files))

	if not pdf_src_files and pdf_src_dir:
		log("Читаю файлы ...")
		pdf_src_files = listdir(longpath(pdf_src_dir), '*.pdf')
		log("%d файлов PDF." % len(pdf_src_files))


	# пересекаем имена файлов без расширений (stem)
	docx_set = set(cut_dir_ext(docx_src_files))
	pdf_set  = set(cut_dir_ext(pdf_src_files))

	common = docx_set & pdf_set
	log("%d имён документов совпадает." % len(common))

	only_docx = docx_set - pdf_set
	log("Документов DOCX без пары PDF: %d" % len(only_docx))

	if only_docx:
		p = log_dir / 'DOCX_без_пары.txt'
		p.write_text('\n'.join(sorted(only_docx)))
		log("	Список сохранён в: " + strip_path_prefix(p))


	only_pdf = pdf_set - docx_set
	log("Документов PDF без пары DOCX: %d" % len(only_pdf))

	if only_pdf:
		p = log_dir / 'PDF_без_пары.txt'
		p.write_text('\n'.join(sorted(only_pdf)))
		log("	Список сохранён в: " + strip_path_prefix(p))


	if not common:
		return ()

	sort_key = lambda p: p.stem
	docx_files = sorted((fp for fp in docx_src_files if fp.stem in common), key=sort_key)
	pdf_files  = sorted((fp for fp in pdf_src_files  if fp.stem in common), key=sort_key)
	assert len(docx_files) == len(pdf_files), (len(docx_files), len(pdf_files))

	pairs = [jsObj(docx=docx, pdf=pdf)
		for docx, pdf in zip(docx_files, pdf_files)
	]

	log("%d пар документов подготовлено к обработке." % len(common))
	log("      ======")

	return pairs


def process_docx_pdf(config: dict={}):
	config = jsObj(config)
	log = lambda *a: _log(*a, config.progress_callback)

	log_dir  = create_dir(config.log_dir  or '.')
	temp_dir = create_dir(config.temp_dir  or './temp')
	result_dir = create_dir(config.result_dir  or './result')

	tasks_total = len(config.document_pairs)

	i = -1  # счётчик может быть не инициализирован циклом
	for i, pair in enumerate(config.document_pairs):
		name = pair.pdf.stem
		result_pdf_path = result_dir / pair.pdf.name
		if result_pdf_path.exists():
			log("Уже готово: %s" % name, i, tasks_total)
			continue

		tmp_pdf_path = temp_dir / pair.pdf.name

		if tmp_pdf_path.exists():
			log("DOCX уже сконвертирован: %s" % name, i, tasks_total)
		else:
			log("Начинаю конвертацию с помощью MS WORD: %s" % name, i, tasks_total)
			try:
				convert(str(pair.docx), str(tmp_pdf_path), keep_active=True)
			except Exception as e:
				log('ERROR')
				log('Ошибка при работе с MS Word:')
				log('\t' + e.__class__.__name__)
				log('\t' + str(e))
				log('Советы:')
				log(' 1. Убедитесь, что Microsoft Word установлен и работает нормально.')
				log(' 2. Закройте все обрабатываемые документы, если они открыты в Word.')
				log(' 3. Разместите файлы DOCX по более короткому пути (напр., ближе к диску C:\\).')
				log(' 4. Не работайте с Word, пока идёт обработка документов.')
				log_error_to_file("Не получилось сконвертировать DOCX в PDF с помощью MS WORD: %s" % strip_path_prefix(pair.docx) + '\n\t\t' + e.__class__.__name__ + ':\t' + str(e), log_dir)
				if config.stop_on_error:
					break
				continue

		if not tmp_pdf_path.exists():
			log('ERROR')
			log("Не получилось сконвертировать DOCX в PDF: %s" % name)
			if config.stop_on_error:
				log('Попробуйте ещё раз.')
				log_error_to_file("Не получилось сконвертировать DOCX в PDF с помощью MS WORD: %s" % strip_path_prefix(pair.docx), log_dir)
				break
			continue

		log("Начинаю сведение двух PDF в один: %s" % name, i, tasks_total)
		# читаем оба PDF
		try:
			pdf_from_docx = open_pdf_Document(tmp_pdf_path, log, log_dir)
			pdf_from_pdf  = open_pdf_Document(pair.pdf, 	log, log_dir)
		except Exception as e:
			if config.stop_on_error:
				break
			continue

		success = mix_pdfs(pdf_from_docx, pdf_from_pdf, result_pdf_path, log, config)
		if not success and config.stop_on_error:
			break

	log("      ======")
	log("Завершено (сделано попыток обработать файлы: %d)." % (i + 1))


def open_pdf_Document(path, log=None, log_dir=Path('.')):
	try:
		with open(path, "rb") as pdf_file_handle:
			document = PDF.loads(pdf_file_handle)
		return document
	except Exception as e:
		log('ERROR')
		log('Ошибка при чтении файла: %s' % strip_path_prefix(path))
		# log('\t' + e.__class__.__name__)
		# log('\t' + str(e))
		log('Совет:')
		log(' - Если это файл открыт в другом приложении, закройте его.')
		log_error_to_file('Ошибка при чтении файла PDF: '+strip_path_prefix(path)+'\n\t\t' + e.__class__.__name__ + ':\t' + str(e), log_dir)
		raise e


def mix_pdfs(pdf_from_docx, pdf_from_pdf, result_pdf_path, log, config):
	try:
		pages_docx = int(pdf_from_docx.get_document_info().get_number_of_pages())
		pages_pdf  = int(pdf_from_pdf .get_document_info().get_number_of_pages())
	except Exception as e:
		log('ERROR')
		log('Проблемы с получением числа страниц PDF.')
		log('\t' + e.__class__.__name__)
		log('\t' + str(e))
		return False

	pdf_end_page = config.get_pages_from_pdf or 3
	docx_start_page = pdf_end_page

	if pages_docx < pdf_end_page:
		log("WARN")
		log(f"Ошибка содержимого DOCX: число страниц ({pages_docx}) меньше, чем нужно вырезать ({pdf_end_page})!")
		log_error_to_file(f"Ошибка содержимого DOCX: число страниц ({pages_docx}) меньше, чем нужно вырезать ({pdf_end_page})!  Имя: " +str(result_pdf_path.stem), create_dir(config.log_dir))
		return False

	if pages_pdf < pdf_end_page:
		log("WARN")
		log(f"Ошибка содержимого PDF: число страниц ({pages_pdf}) меньше, чем нужно ({pdf_end_page})!")
		log_error_to_file(f"Ошибка содержимого PDF: число страниц ({pages_pdf}) меньше, чем нужно ({pdf_end_page})!  Имя: " +str(result_pdf_path.stem), create_dir(config.log_dir))
		return False

	if pages_docx != pages_pdf:
		log("WARN")
		log(f"Предупреждение: число страниц из PDF ({pages_pdf}) и DOCX ({pages_docx}) различается!")
		log_warning_to_file(f"Число страниц из PDF ({pages_pdf}) и DOCX ({pages_docx}) различается!  Выравнивание по концу документа: {'в' if config.smart_solution_on_different_sizes else 'от'}ключено.  Имя: " +str(result_pdf_path.stem), create_dir(config.log_dir))

		if config.smart_solution_on_different_sizes:
			if pages_pdf - pages_docx > pdf_end_page:
				log("  Размеры документов различаются слишком сильно, ")
				log("    выровнять их невозможно.")
				log("  Пожалуйста, проверьте эту пару документов.")
				log_error_to_file("Размеры документов различаются слишком сильно, выровнять их невозможно.  Имя: " +str(result_pdf_path.stem), create_dir(config.log_dir))
				return False
			log("  Применяю выравнивание по концу документа:")
			docx_start_page = pages_docx - (pages_pdf - pdf_end_page)
			log("    Из DOCX будут взяты страницы: с %d по %d (до конца)." % (docx_start_page + 1, pages_docx))

	# Создаём пустой PDF-файл
	output_pdf = Document()

	# Сведение
	for i in range(0, pdf_end_page):
		output_pdf.append_page(pdf_from_pdf.get_page(i))
	for i in range(docx_start_page, pages_docx):
		output_pdf.append_page(pdf_from_docx.get_page(i))

	try:
		# Запись PDF
		with open(result_pdf_path, "wb") as pdf_out_handle:
			PDF.dumps(pdf_out_handle, output_pdf)
	except Exception as e:
		log('ERROR')
		log('Проблемы с записью итогового PDF на диск:')
		log('\t' + e.__class__.__name__)
		log('\t' + str(e))
		log('Пожалуйста, проверьте, что этот файл нигде не открыт (для просмотра):')
		log('\t' + str(result_pdf_path))
		log_error_to_file('Проблемы с записью итогового PDF на диск: ' +strip_path_prefix(result_pdf_path), create_dir(config.log_dir))
		return False

	return True


def read_yml(file_path):
	try:
		with open(file_path) as f:
			data = yaml.load(f, Loader=yaml.FullLoader)
		return data
	except Exception as e:
		print('ERROR')
		print('НЕ МОГУ ПРОЧИТАТЬ ФАЙЛ КОНФИГУРАЦИИ:')
		print('\t' + e.__class__.__name__)
		print('\t' + str(e))
		print('Файл конфигурации должен находиться здесь:')
		print("\t" + file_path)
		print("\tАбсолютный путь:  ",Path(file_path).absolute())
		exit(2)


def console_progress(msg, i=None, total=None):
	print(strftime("[%y-%m-%d %H:%M:%S]", localtime()), end='  ')
	if total:
		total_str = str(total)
		print(f"%{len(total_str)}d/%s" % (i, total_str), end='  ')
	print(msg)


def log_error_to_file(msg:str, dir_path:Path):
	with (dir_path / ERRORS_FILE).open('a') as f:
		f.write(strftime("[%Y-%m-%d %H:%M:%S]", localtime()) + '  ')
		f.write(msg)
		f.write('\n')


def log_warning_to_file(msg:str, dir_path:Path):
	with (dir_path / WARNINGS_FILE).open('a') as f:
		f.write(strftime("[%Y-%m-%d %H:%M:%S]", localtime()) + '  ')
		f.write(msg)
		f.write('\n')



def main(config_file_path=CONFIG_FILE):
	config = read_yml(config_file_path)
	config['progress_callback'] = console_progress

	try:
		pairs = find_files(config)
	except Exception as e:
		print('ERROR searching files:')
		print('\t' + e.__class__.__name__)
		print('\t' + str(e))
		exit(1)

	config['document_pairs'] = pairs
	try:
		process_docx_pdf(config)
	except Exception as e:
		print('ERROR processing documents:')
		print('\t' + e.__class__.__name__)
		print('\t' + str(e))
		exit(1)



if __name__ == '__main__':
	main()
