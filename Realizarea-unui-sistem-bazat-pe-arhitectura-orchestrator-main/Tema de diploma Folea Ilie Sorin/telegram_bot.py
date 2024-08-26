import os
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes


TOKEN = ''

# Partitia de baza unde se vor cauta fișierele
BASE_DIR = r'C:\\'

# Variabila globala pentru a controla starea botului
bot_active = False

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    global bot_active
    bot_active = True

    keyboard = [['/start', '/stop'], ['/help', '/listdir', '/search']]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True)

    await update.message.reply_text(
        'Salut! Trimite-mi numele fișierului pe care vrei să-l primești.',
        reply_markup=reply_markup
    )

async def stop(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    global bot_active
    bot_active = False
    await update.message.reply_text('Botul a fost oprit.')

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    help_text = (
        "/start - Pornește botul\n"
        "/stop - Oprește botul\n"
        "/help - Afișează acest mesaj de ajutor\n"
        "/listdir [cale] - Listează toate fișierele dintr-un director specificat\n"
        "/search [pattern] - Caută fișiere după un model\n"
        "Trimite numele unui fișier pentru a-l primi."
    )
    await update.message.reply_text(help_text)

async def list_files_in_directory(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not bot_active:
        await update.message.reply_text('Botul nu este activ. Te rog să trimiți comanda /start pentru a activa botul.')
        return

    directory = update.message.text.split(' ', 1)[1] if ' ' in update.message.text else BASE_DIR

    if not os.path.isdir(directory):
        await update.message.reply_text(f'Directorul specificat nu există: {directory}')
        return

    files_list = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            files_list.append(os.path.join(root, file))
        break  # Nu parcurge recursiv pentru a evita listarea fișierelor din subdirectoare

    if files_list:
        chunk_size = 50
        for i in range(0, len(files_list), chunk_size):
            await update.message.reply_text('\n'.join(files_list[i:i + chunk_size]))
    else:
        await update.message.reply_text('Nu s-au găsit fișiere în directorul specificat.')

async def search_files(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not bot_active:
        await update.message.reply_text('Botul nu este activ. Te rog să trimiți comanda /start pentru a activa botul.')
        return

    pattern = update.message.text.split(' ', 1)[1] if ' ' in update.message.text else '*'
    files_list = []
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if fnmatch.fnmatch(file, pattern):
                files_list.append(os.path.join(root, file))

    if files_list:
        chunk_size = 50
        for i in range(0, len(files_list), chunk_size):
            await update.message.reply_text('\n'.join(files_list[i:i + chunk_size]))
    else:
        await update.message.reply_text('Nu s-au găsit fișiere care să corespundă modelului specificat.')

def find_file(file_name):
    for root, dirs, files in os.walk(BASE_DIR):
        if file_name in files:
            return os.path.join(root, file_name)
    return None

async def send_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not bot_active:
        await update.message.reply_text('Botul nu este activ. Te rog să trimiți comanda /start pentru a activa botul.')
        return

    file_name = update.message.text
    chat_id = update.message.chat_id

    file_path = find_file(file_name)

    if file_path:
        await context.bot.send_document(chat_id=chat_id, document=open(file_path, 'rb'))
    else:
        await update.message.reply_text(f'Fișierul nu există: {file_name}')

def main():
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("stop", stop))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("listdir", list_files_in_directory))
    application.add_handler(CommandHandler("search", search_files))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, send_file))

    application.run_polling()

if __name__ == '__main__':
    main()
