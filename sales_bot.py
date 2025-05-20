import logging
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes, ConversationHandler, MessageHandler, filters
from telegram.error import Conflict, NetworkError, BadRequest, TelegramError
import openpyxl
from openpyxl import Workbook
import os
import sys
from pathlib import Path
import time
import asyncio

# Enable logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    handlers=[
        logging.FileHandler('bot.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# States for conversation
CHOOSING, CLASSIFYING, GETTING_INFO, PRODUCT_TYPE, BUDGET, TIMELINE, COMPANY_INFO = range(7)

# Customer classifications
CUSTOMER_TYPES = {
    'hot': 'عميل محتمل عالي',
    'warm': 'عميل محتمل متوسط',
    'cold': 'عميل محتمل منخفض'
}

# Product types
PRODUCT_TYPES = {
    'software': 'برامج إدارة المبيعات',
    'hardware': 'أجهزة ومعدات',
    'service': 'خدمات استشارية',
    'other': 'أخرى'
}

# Excel file setup
EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'customer_data.xlsx')

def setup_excel():
    """Create or load Excel file with headers if it doesn't exist."""
    try:
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            ws = wb.active
            headers = [
                'التاريخ',
                'اسم المستخدم',
                'معرف المستخدم',
                'نوع العميل',
                'رقم الهاتف',
                'البريد الإلكتروني',
                'نوع المنتج',
                'الميزانية المتوقعة',
                'موعد الشراء المتوقع',
                'اسم الشركة',
                'حجم الشركة',
                'ملاحظات إضافية'
            ]
            ws.append(headers)
            wb.save(EXCEL_FILE)
            logging.info(f"Created new Excel file at {EXCEL_FILE}")
    except PermissionError:
        logging.error(f"Permission denied when creating Excel file at {EXCEL_FILE}")
        print(f"خطأ: لا يمكن الوصول إلى الملف {EXCEL_FILE}. تأكد من إغلاق الملف إذا كان مفتوحاً.")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error creating Excel file: {str(e)}")
        print(f"حدث خطأ أثناء إنشاء ملف Excel: {str(e)}")
        sys.exit(1)

def save_to_excel(user_data):
    """Save customer data to Excel file."""
    try:
        if not os.path.exists(EXCEL_FILE):
            setup_excel()
            
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        
        # Prepare data row
        data = [
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            user_data.get('username', ''),
            user_data.get('user_id', ''),
            user_data.get('customer_type', ''),
            user_data.get('phone', ''),
            user_data.get('email', ''),
            user_data.get('product_type', ''),
            user_data.get('budget', ''),
            user_data.get('timeline', ''),
            user_data.get('company_name', ''),
            user_data.get('company_size', ''),
            user_data.get('notes', '')
        ]
        
        ws.append(data)
        wb.save(EXCEL_FILE)
        logging.info("Successfully saved customer data to Excel")
    except PermissionError:
        logging.error(f"Permission denied when saving to Excel file at {EXCEL_FILE}")
        print(f"خطأ: لا يمكن حفظ البيانات في الملف {EXCEL_FILE}. تأكد من إغلاق الملف إذا كان مفتوحاً.")
    except Exception as e:
        logging.error(f"Error saving to Excel: {str(e)}")
        print(f"حدث خطأ أثناء حفظ البيانات: {str(e)}")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /help is issued."""
    try:
        logger.info(f"Help command received from user {update.effective_user.id}")
        await update.message.reply_text(
            "مرحباً بك في بوت المبيعات!\n\n"
            "الأوامر المتاحة:\n"
            "/start - بدء محادثة جديدة\n"
            "/help - عرض هذه الرسالة\n"
            "/cancel - إلغاء المحادثة الحالية"
        )
    except Exception as e:
        logger.error(f"Error in help command: {e}")
        await update.message.reply_text("عذراً، حدث خطأ. يرجى المحاولة مرة أخرى.")

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancel and end the conversation."""
    try:
        logger.info(f"Cancel command received from user {update.effective_user.id}")
        await update.message.reply_text(
            "تم إلغاء المحادثة. يمكنك البدء من جديد باستخدام الأمر /start"
        )
        return ConversationHandler.END
    except Exception as e:
        logger.error(f"Error in cancel command: {e}")
        return ConversationHandler.END

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Start the conversation and ask user about their interest."""
    try:
        logger.info(f"Start command received from user {update.effective_user.id}")
        
        # Clear any existing conversation data
        context.user_data.clear()
        
        # Store user info
        context.user_data['username'] = update.effective_user.username or update.effective_user.first_name
        context.user_data['user_id'] = update.effective_user.id
        
        keyboard = [
            [
                InlineKeyboardButton("نعم، أنا مهتم", callback_data='interested'),
                InlineKeyboardButton("لا، لست مهتماً", callback_data='not_interested')
            ]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await update.message.reply_text(
            "مرحباً! أنا بوت المبيعات الخاص ببرنامج البدر.\n"
            "هل أنت مهتم بمنتجاتنا؟",
            reply_markup=reply_markup
        )
        logger.info(f"Start message sent to user {update.effective_user.id}")
        return CHOOSING
    except Exception as e:
        logger.error(f"Error in start command: {e}")
        await update.message.reply_text(
            "عذراً، حدث خطأ في بدء المحادثة. يرجى المحاولة مرة أخرى باستخدام /start"
        )
        return ConversationHandler.END

async def button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle button presses."""
    query = update.callback_query
    await query.answer()
    
    if query.data == 'interested':
        keyboard = [
            [
                InlineKeyboardButton("عميل محتمل عالي", callback_data='hot'),
                InlineKeyboardButton("عميل محتمل متوسط", callback_data='warm')
            ],
            [
                InlineKeyboardButton("عميل محتمل منخفض", callback_data='cold')
            ]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await query.edit_message_text(
            text="ممتاز! كيف تقيم مستوى اهتمامك بمنتجاتنا؟",
            reply_markup=reply_markup
        )
        return CLASSIFYING
    
    elif query.data in CUSTOMER_TYPES:
        context.user_data['customer_type'] = CUSTOMER_TYPES[query.data]
        await query.edit_message_text(
            "من فضلك، قم بإدخال رقم هاتفك:"
        )
        return GETTING_INFO
    
    elif query.data in PRODUCT_TYPES:
        context.user_data['product_type'] = PRODUCT_TYPES[query.data]
        await query.edit_message_text(
            "ما هي الميزانية المتوقعة للمشروع؟ (بالريال السعودي)"
        )
        return BUDGET
    
    else:
        await query.edit_message_text(
            "شكراً لوقتك! إذا كنت مهتماً في المستقبل، لا تتردد في التواصل معنا."
        )
        return ConversationHandler.END

async def get_contact_info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Get customer contact information."""
    text = update.message.text
    
    if not context.user_data.get('phone'):
        context.user_data['phone'] = text
        await update.message.reply_text(
            "شكراً! من فضلك، قم بإدخال بريدك الإلكتروني:"
        )
        return GETTING_INFO
    else:
        context.user_data['email'] = text
        keyboard = [
            [
                InlineKeyboardButton("برامج إدارة المبيعات", callback_data='software'),
                InlineKeyboardButton("أجهزة ومعدات", callback_data='hardware')
            ],
            [
                InlineKeyboardButton("خدمات استشارية", callback_data='service'),
                InlineKeyboardButton("أخرى", callback_data='other')
            ]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text(
            "ما هو نوع المنتج الذي تهتم به؟",
            reply_markup=reply_markup
        )
        return PRODUCT_TYPE

async def get_budget(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Get customer budget information."""
    context.user_data['budget'] = update.message.text
    await update.message.reply_text(
        "ما هو موعد الشراء المتوقع؟ (مثال: خلال شهر، خلال 3 أشهر، خلال سنة)"
    )
    return TIMELINE

async def get_timeline(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Get customer timeline information."""
    context.user_data['timeline'] = update.message.text
    await update.message.reply_text(
        "ما هو اسم شركتك؟"
    )
    return COMPANY_INFO

async def get_company_info(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Get company information."""
    text = update.message.text.lower()
    
    if not context.user_data.get('company_name'):
        context.user_data['company_name'] = update.message.text
        await update.message.reply_text(
            "ما هو حجم شركتك؟ (عدد الموظفين)"
        )
        return COMPANY_INFO
    elif not context.user_data.get('company_size'):
        context.user_data['company_size'] = update.message.text
        await update.message.reply_text(
            "هل لديك أي ملاحظات أو متطلبات إضافية؟\n"
            "يمكنك كتابة 'لا' أو 'لا شكراً' إذا لم يكن لديك ملاحظات."
        )
        return COMPANY_INFO
    else:
        # Handle the notes response
        if text in ['لا', 'لا شكرا', 'لا شكراً', 'لا يوجد', 'لا شيء']:
            context.user_data['notes'] = 'لا توجد ملاحظات'
        else:
            context.user_data['notes'] = update.message.text
            
        # Save all data to Excel
        save_to_excel(context.user_data)
        
        await update.message.reply_text(
            f"شكراً لك! تم تصنيفك كـ {context.user_data['customer_type']}.\n"
            "تم حفظ جميع معلوماتك وسيتواصل معك فريق المبيعات قريباً."
        )
        return ConversationHandler.END

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle any message that is not a command."""
    try:
        logger.info(f"Message received from user {update.effective_user.id}: {update.message.text}")
        await update.message.reply_text(
            "عذراً، لا أستطيع فهم هذه الرسالة.\n"
            "استخدم /start لبدء محادثة جديدة أو /help للحصول على المساعدة."
        )
    except Exception as e:
        logger.error(f"Error in message handler: {e}")

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle errors in the bot."""
    logger.error(f"Update {update} caused error {context.error}")
    if isinstance(context.error, Conflict):
        print("خطأ: يبدو أن هناك نسخة أخرى من البوت تعمل بالفعل.")
        print("يرجى إغلاق جميع نسخ البوت الأخرى وإعادة المحاولة.")
        sys.exit(1)
    elif isinstance(context.error, NetworkError):
        print("خطأ في الاتصال بالإنترنت. جاري إعادة المحاولة...")
        time.sleep(5)
    else:
        print(f"حدث خطأ غير متوقع: {context.error}")

def main() -> None:
    """Start the bot."""
    try:
        # Setup Excel file
        setup_excel()
        
        # Create the Application and pass it your bot's token
        application = Application.builder().token('7721450926:AAHc3xILNUT4JqRiQJs0WF3gSH9j3QdUndU').build()

        # Add error handler
        application.add_error_handler(error_handler)

        # Add conversation handler
        conv_handler = ConversationHandler(
            entry_points=[CommandHandler('start', start)],
            states={
                CHOOSING: [CallbackQueryHandler(button)],
                CLASSIFYING: [CallbackQueryHandler(button)],
                GETTING_INFO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_contact_info)],
                PRODUCT_TYPE: [CallbackQueryHandler(button)],
                BUDGET: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_budget)],
                TIMELINE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_timeline)],
                COMPANY_INFO: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_company_info)],
            },
            fallbacks=[],
        )

        application.add_handler(conv_handler)

        print("جاري تشغيل البوت...")
        print("اضغط Ctrl+C للإيقاف")
        
        # Start the Bot
        application.run_polling(allowed_updates=Update.ALL_TYPES)
        
    except Conflict:
        print("خطأ: يبدو أن هناك نسخة أخرى من البوت تعمل بالفعل.")
        print("يرجى إغلاق جميع نسخ البوت الأخرى وإعادة المحاولة.")
        sys.exit(1)
    except Exception as e:
        print(f"حدث خطأ غير متوقع: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()