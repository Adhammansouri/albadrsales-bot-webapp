import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import arabic_reshaper
from bidi.algorithm import get_display
import warnings
warnings.filterwarnings('ignore')

class CustomerAnalyzer:
    def __init__(self, excel_file='customer_data.xlsx'):
        self.excel_file = excel_file
        self.df = None
        self.load_data()
        
    def load_data(self):
        """تحميل البيانات من ملف Excel"""
        try:
            self.df = pd.read_excel(self.excel_file)
            # حذف الأعمدة الفارغة
            self.df = self.df.loc[:, ~self.df.columns.str.contains('^Unnamed')]
            print("\nالأعمدة المتاحة في الملف:")
            print(self.df.columns.tolist())
            print("\nتم تحميل البيانات بنجاح!")
        except Exception as e:
            print(f"حدث خطأ أثناء تحميل البيانات: {str(e)}")
            return None

    def analyze_customer_types(self):
        """تحليل أنواع العملاء"""
        if self.df is None:
            return
        
        try:
            customer_types = self.df['نوع العميل'].value_counts()
            print("\nتحليل أنواع العملاء:")
            print("-------------------")
            for customer_type, count in customer_types.items():
                print(f"{customer_type}: {count} عميل")
            
            # رسم بياني دائري
            plt.figure(figsize=(10, 6))
            plt.pie(customer_types.values, labels=customer_types.index, autopct='%1.1f%%')
            plt.title('توزيع أنواع العملاء')
            plt.savefig('customer_types.png')
            plt.close()
        except KeyError:
            print("لم يتم العثور على عمود 'نوع العميل' في البيانات")
        except Exception as e:
            print(f"حدث خطأ أثناء تحليل أنواع العملاء: {str(e)}")

    def analyze_contact_info(self):
        """تحليل معلومات الاتصال"""
        if self.df is None:
            return
        
        try:
            print("\nتحليل معلومات الاتصال:")
            print("---------------------")
            
            # تحليل البريد الإلكتروني
            email_count = self.df['البريد الإلكتروني'].notna().sum()
            print(f"عدد العملاء الذين لديهم بريد إلكتروني: {email_count}")
            
            # تحليل رقم الهاتف
            phone_count = self.df['رقم الهاتف'].notna().sum()
            print(f"عدد العملاء الذين لديهم رقم هاتف: {phone_count}")
            
            # رسم بياني للمعلومات المتوفرة
            contact_data = {
                'بريد إلكتروني': email_count,
                'رقم هاتف': phone_count
            }
            
            plt.figure(figsize=(10, 6))
            plt.bar(contact_data.keys(), contact_data.values())
            plt.title('توفر معلومات الاتصال')
            plt.ylabel('عدد العملاء')
            plt.savefig('contact_info.png')
            plt.close()
            
        except Exception as e:
            print(f"حدث خطأ أثناء تحليل معلومات الاتصال: {str(e)}")

    def analyze_temporal_data(self):
        """تحليل البيانات الزمنية"""
        if self.df is None:
            return
        
        try:
            # تحويل عمود التاريخ إلى datetime
            self.df['التاريخ'] = pd.to_datetime(self.df['التاريخ'])
            
            # تحليل حسب اليوم
            daily_counts = self.df['التاريخ'].dt.date.value_counts().sort_index()
            
            print("\nتحليل البيانات الزمنية:")
            print("---------------------")
            print("عدد العملاء حسب اليوم:")
            for date, count in daily_counts.items():
                print(f"{date}: {count} عميل")
            
            # رسم بياني للاتجاه الزمني
            plt.figure(figsize=(12, 6))
            plt.plot(daily_counts.index, daily_counts.values, marker='o')
            plt.title('اتجاه تسجيل العملاء')
            plt.xlabel('التاريخ')
            plt.ylabel('عدد العملاء')
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig('temporal_analysis.png')
            plt.close()
            
        except Exception as e:
            print(f"حدث خطأ أثناء تحليل البيانات الزمنية: {str(e)}")

    def generate_sales_recommendations(self):
        """توليد توصيات للمبيعات"""
        if self.df is None:
            return
        
        try:
            print("\nتوصيات المبيعات:")
            print("---------------")
            
            # تحليل العملاء المحتملين العاليين
            hot_customers = self.df[self.df['نوع العميل'] == 'عميل محتمل عالي']
            if not hot_customers.empty:
                print("\n1. العملاء المحتملين العاليين:")
                for _, customer in hot_customers.iterrows():
                    print(f"- {customer['اسم المستخدم']}")
                    if pd.notna(customer['البريد الإلكتروني']):
                        print(f"  البريد الإلكتروني: {customer['البريد الإلكتروني']}")
                    if pd.notna(customer['رقم الهاتف']):
                        print(f"  رقم الهاتف: {customer['رقم الهاتف']}")
            
            # تحليل العملاء حسب نوع العميل
            customer_types = self.df['نوع العميل'].value_counts()
            print("\n2. توزيع العملاء حسب النوع:")
            for customer_type, count in customer_types.items():
                print(f"- {customer_type}: {count} عميل")
            
            # تحليل العملاء حسب التاريخ
            recent_customers = self.df.sort_values('التاريخ', ascending=False).head(3)
            print("\n3. أحدث العملاء:")
            for _, customer in recent_customers.iterrows():
                print(f"- {customer['اسم المستخدم']} ({customer['التاريخ'].strftime('%Y-%m-%d')})")
                print(f"  نوع العميل: {customer['نوع العميل']}")
                
        except Exception as e:
            print(f"حدث خطأ أثناء توليد توصيات المبيعات: {str(e)}")

    def generate_report(self):
        """توليد تقرير شامل"""
        print("\nتقرير تحليل العملاء")
        print("==================")
        
        self.analyze_customer_types()
        self.analyze_contact_info()
        self.analyze_temporal_data()
        self.generate_sales_recommendations()
        
        print("\nتم حفظ الرسوم البيانية في الملفات التالية:")
        print("- customer_types.png: توزيع أنواع العملاء")
        print("- contact_info.png: توفر معلومات الاتصال")
        print("- temporal_analysis.png: اتجاه تسجيل العملاء")

def main():
    analyzer = CustomerAnalyzer()
    analyzer.generate_report()

if __name__ == "__main__":
    main() 