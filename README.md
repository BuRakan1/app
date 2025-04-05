import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import sqlite3
import pandas as pd
from tkcalendar import DateEntry
import hashlib
import re
import json
import pyperclip
import zipfile
import tempfile
import shutil
import datetime
import hardware_verification
import uuid
import project_structure
import version_manager

# إعداد هيكل المشروع
project_dirs = project_structure.setup_project_structure()

# طباعة معلومات عن مسارات المشروع (للتصحيح أثناء التطوير)
print("مجلد التطبيق:", project_dirs["base_dir"])
print("مجلد الموارد:", project_dirs["resources_dir"])
print("مجلد البيانات:", project_dirs["data_dir"])
print("مسار قاعدة البيانات:", project_dirs["database_path"])

class ArchiveManager:
    def __init__(self, root, main_app, colors, fonts):
        self.root = root
        self.main_app = main_app
        self.colors = colors
        self.fonts = fonts
        self.conn = None
        self.archive_data = None
        self.current_archive_path = None

    def export_courses_to_archive(self, course_names):
        """تصدير دورة أو عدة دورات إلى ملف أرشيف مع إضافة فئة الدورة"""
        if not course_names:
            return False

        # إنشاء مسار ملف الأرشيف
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        archive_name = "course_archive_" + timestamp

        if len(course_names) == 1:
            archive_name = course_names[0].replace(" ", "_") + "_" + timestamp

        export_file = filedialog.asksaveasfilename(
            defaultextension=".crsarch",
            filetypes=[("ملفات أرشيف الدورات", "*.crsarch")],
            initialfile=archive_name
        )

        if not export_file:
            return False

        # تحديد فئة كل دورة قبل التصدير
        course_categories = {}

        # إنشاء نافذة تحديد الفئات
        category_window = tk.Toplevel(self.root)
        category_window.title("تحديد فئة الدورة")
        category_window.geometry("500x400")
        category_window.configure(bg=self.colors["light"])
        category_window.transient(self.root)
        category_window.grab_set()

        # توسيط النافذة
        x = (category_window.winfo_screenwidth() - 500) // 2
        y = (category_window.winfo_screenheight() - 400) // 2
        category_window.geometry(f"500x400+{x}+{y}")

        tk.Label(
            category_window,
            text="تحديد فئة الدورات",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        tk.Label(
            category_window,
            text="الرجاء تحديد فئة كل دورة:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            pady=10
        ).pack()

        # القائمة المنسدلة للفئات
        categories = ["ضباط", "أفراد", "مشتركة", "مدنيين", "طلبة"]

        # إنشاء إطار لكل دورة
        courses_frame = tk.Frame(category_window, bg=self.colors["light"])
        courses_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        category_vars = {}

        for i, course in enumerate(course_names):
            course_frame = tk.Frame(courses_frame, bg=self.colors["light"], pady=5)
            course_frame.pack(fill=tk.X)

            tk.Label(
                course_frame,
                text=f"دورة: {course}",
                font=self.fonts["text_bold"],
                bg=self.colors["light"],
                width=20,
                anchor=tk.E
            ).pack(side=tk.RIGHT, padx=5)

            category_var = tk.StringVar(value=categories[0])
            category_vars[course] = category_var

            category_dropdown = ttk.Combobox(
                course_frame,
                textvariable=category_var,
                values=categories,
                state="readonly",
                width=15,
                font=self.fonts["text"]
            )
            category_dropdown.pack(side=tk.RIGHT, padx=5)

        # متغير للتحقق من اكتمال العملية
        completed = [False]

        def confirm_categories():
            for course in course_names:
                course_categories[course] = category_vars[course].get()
            completed[0] = True
            category_window.destroy()

        # أزرار التأكيد والإلغاء
        button_frame = tk.Frame(category_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=20)

        confirm_btn = tk.Button(
            button_frame,
            text="تأكيد",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=confirm_categories
        )
        confirm_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = tk.Button(
            button_frame,
            text="إلغاء",
            font=self.fonts["text_bold"],
            bg=self.colors["danger"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=category_window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)

        # انتظار حتى يتم إغلاق النافذة
        self.root.wait_window(category_window)

        if not completed[0]:
            return False  # تم إلغاء العملية

        # إنشاء نافذة تقدم العملية
        progress_window = tk.Toplevel(self.root)
        progress_window.title("تصدير الدورات إلى الأرشيف")
        progress_window.geometry("450x180")
        progress_window.configure(bg=self.colors["light"])
        progress_window.transient(self.root)
        progress_window.grab_set()

        # توسيط النافذة
        x = (progress_window.winfo_screenwidth() - 450) // 2
        y = (progress_window.winfo_screenheight() - 180) // 2
        progress_window.geometry(f"450x180+{x}+{y}")

        tk.Label(
            progress_window,
            text=f"جاري تصدير {len(course_names)} دورة إلى الأرشيف...",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            pady=10
        ).pack()

        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(
            progress_window,
            variable=progress_var,
            maximum=100,
            length=400
        )
        progress_bar.pack(pady=10)

        status_label = tk.Label(
            progress_window,
            text="جاري تحضير البيانات...",
            font=self.fonts["text"],
            bg=self.colors["light"]
        )
        status_label.pack(pady=5)

        progress_window.update()

        try:
            # إنشاء مجلد مؤقت للتصدير
            temp_dir = tempfile.mkdtemp()

            # إنشاء قاموس البيانات الأساسي
            archive_data = {
                "metadata": {
                    "creation_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "courses_count": len(course_names),
                    "course_names": course_names,
                    "archive_id": str(uuid.uuid4())
                },
                "courses": {}
            }

            # استخراج البيانات لكل دورة
            cursor = self.main_app.conn.cursor()

            for i, course_name in enumerate(course_names):
                # تحديث شريط التقدم
                progress_var.set((i / len(course_names)) * 40)
                status_label.config(text=f"جاري استخراج بيانات دورة: {course_name}")
                progress_window.update()

                course_data = {
                    "course_info": {
                        "name": course_name,
                        "category": course_categories[course_name]  # إضافة فئة الدورة
                    },
                    "students": [],
                    "attendance": [],
                    "sections": [],
                    "student_sections": []
                }

                # استخراج بيانات الطلاب
                cursor.execute("SELECT * FROM trainees WHERE course=?", (course_name,))
                students = cursor.fetchall()

                # تحويل بيانات الطلاب إلى قائمة من القواميس
                column_names = [description[0] for description in cursor.description]
                for student in students:
                    student_dict = {column_names[i]: student[i] for i in range(len(column_names))}
                    course_data["students"].append(student_dict)

                    # استخراج سجلات الحضور لكل طالب
                    cursor.execute("SELECT * FROM attendance WHERE national_id=?", (student[0],))
                    attendance_records = cursor.fetchall()

                    att_column_names = [description[0] for description in cursor.description]
                    for record in attendance_records:
                        record_dict = {att_column_names[i]: record[i] for i in range(len(att_column_names))}
                        course_data["attendance"].append(record_dict)

                # استخراج بيانات الفصول
                cursor.execute("SELECT * FROM course_sections WHERE course_name=?", (course_name,))
                sections = cursor.fetchall()

                if sections:
                    section_column_names = [description[0] for description in cursor.description]
                    for section in sections:
                        section_dict = {section_column_names[i]: section[i] for i in range(len(section_column_names))}
                        course_data["sections"].append(section_dict)

                        # استخراج توزيع الطلاب على الفصول
                        cursor.execute("""
                            SELECT * FROM student_sections 
                            WHERE course_name=? AND section_name=?
                        """, (course_name, section[2]))

                        student_sections = cursor.fetchall()

                        if student_sections:
                            ss_column_names = [description[0] for description in cursor.description]
                            for ss in student_sections:
                                ss_dict = {ss_column_names[i]: ss[i] for i in range(len(ss_column_names))}
                                course_data["student_sections"].append(ss_dict)

                # حفظ بيانات الدورة
                archive_data["courses"][course_name] = course_data

                # تحديث شريط التقدم
                progress_var.set(40 + (i / len(course_names)) * 30)
                progress_window.update()

            # حفظ البيانات في ملف JSON
            progress_var.set(70)
            status_label.config(text="جاري حفظ البيانات...")
            progress_window.update()

            archive_json = os.path.join(temp_dir, "archive_data.json")
            with open(archive_json, 'w', encoding='utf-8') as f:
                json.dump(archive_data, f, ensure_ascii=False, indent=2)

            # إنشاء ملف الأرشيف المضغوط
            progress_var.set(85)
            status_label.config(text="جاري إنشاء ملف الأرشيف...")
            progress_window.update()

            with zipfile.ZipFile(export_file, 'w', compression=zipfile.ZIP_DEFLATED) as archive_zip:
                archive_zip.write(archive_json, arcname="archive_data.json")
                # يمكن إضافة ملفات أخرى هنا إذا لزم الأمر مستقبلاً

            progress_var.set(100)
            status_label.config(text="تم تصدير الدورات إلى الأرشيف بنجاح!")
            progress_window.update()

            # حذف المجلد المؤقت
            shutil.rmtree(temp_dir)

            # عرض رسالة نجاح العملية
            messagebox.showinfo("نجاح", f"تم تصدير {len(course_names)} دورة إلى الأرشيف بنجاح:\n{export_file}")

            # إغلاق نافذة التقدم بعد ثانيتين
            progress_window.after(2000, progress_window.destroy)

            return True

        except Exception as e:
            # حذف المجلد المؤقت في حالة وجود خطأ
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير الدورات: {str(e)}")
            progress_window.destroy()
            return False

    def load_archive(self, archive_path=None):
        """تحميل ملف أرشيف دورات"""
        if archive_path is None:
            archive_path = filedialog.askopenfilename(
                filetypes=[("ملفات أرشيف الدورات", "*.crsarch"), ("جميع الملفات", "*.*")],
                title="اختر ملف أرشيف دورات"
            )

        if not archive_path:
            return False

        try:
            # إنشاء مجلد مؤقت للاستخراج
            temp_dir = tempfile.mkdtemp()

            # استخراج ملفات الأرشيف
            with zipfile.ZipFile(archive_path, 'r') as archive_zip:
                archive_zip.extractall(temp_dir)

            # قراءة ملف البيانات
            archive_json = os.path.join(temp_dir, "archive_data.json")
            with open(archive_json, 'r', encoding='utf-8') as f:
                self.archive_data = json.load(f)

            # حفظ مسار الأرشيف الحالي
            self.current_archive_path = archive_path

            # حذف المجلد المؤقت
            shutil.rmtree(temp_dir)

            return True

        except Exception as e:
            # حذف المجلد المؤقت في حالة وجود خطأ
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

            messagebox.showerror("خطأ", f"حدث خطأ أثناء تحميل الأرشيف: {str(e)}")
            return False

    def open_archive_window(self):
        """فتح نافذة عرض الأرشيف مع إضافة عرض فئة الدورة وإحصائيات الفئات"""
        if not self.load_archive():
            return

        archive_window = tk.Toplevel(self.root)
        archive_window.title("أرشيف الدورات - قراءة فقط")
        archive_window.geometry("1200x700")  # زيادة عرض النافذة لاستيعاب المعلومات الإضافية
        archive_window.configure(bg=self.colors["light"])

        # جعل النافذة قابلة للتمدد
        archive_window.resizable(True, True)

        # توسيط النافذة
        x = (archive_window.winfo_screenwidth() - 1200) // 2
        y = (archive_window.winfo_screenheight() - 700) // 2
        archive_window.geometry(f"1200x700+{x}+{y}")

        # العنوان مع شريط تمييز بلون مختلف للتنبيه على وضع الأرشيف
        header_frame = tk.Frame(archive_window, bg="#FF5722", height=60)  # لون برتقالي للتمييز
        header_frame.pack(fill=tk.X)

        archive_date = self.archive_data["metadata"]["creation_date"]
        courses_count = self.archive_data["metadata"]["courses_count"]

        tk.Label(
            header_frame,
            text=f"أرشيف الدورات المنعقدة في مدينة تدريب الأمن العام بالمنطقة الشرقية",
            font=self.fonts["large_title"],
            bg="#FF5722",
            fg="white"
        ).pack(side=tk.RIGHT, pady=15, padx=20)

        tk.Label(
            header_frame,
            text=f"تاريخ الأرشيف: {archive_date} | عدد الدورات: {courses_count}",
            font=self.fonts["text_bold"],
            bg="#FF5722",
            fg="white"
        ).pack(side=tk.LEFT, pady=15, padx=20)

        # إضافة إطار للإحصائيات العامة
        stats_frame = tk.LabelFrame(
            archive_window,
            text="إحصائيات عامة حسب الفئات",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=10,
            pady=10
        )
        stats_frame.pack(fill=tk.X, padx=10, pady=5)

        # حساب الإحصائيات العامة للفئات
        category_stats = self.calculate_archive_category_stats()

        # إنشاء إطار للإحصائيات
        categories_stats_frame = tk.Frame(stats_frame, bg=self.colors["light"])
        categories_stats_frame.pack(fill=tk.X, pady=5)

        # الألوان للفئات المختلفة
        category_colors = {
            "ضباط": "#1E88E5",  # أزرق
            "أفراد": "#43A047",  # أخضر
            "مشتركة": "#FB8C00",  # برتقالي
            "مدنيين": "#8E24AA",  # بنفسجي
            "طلبة": "#F4511E"  # أحمر برتقالي
        }

        # إنشاء بطاقات للإحصائيات
        for category, stats in category_stats.items():
            category_frame = tk.Frame(categories_stats_frame, bg=self.colors["light"], bd=1, relief=tk.RIDGE, padx=10,
                                      pady=5)
            category_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)

            # لون الفئة
            color = category_colors.get(category, self.colors["primary"])

            # عنوان الفئة
            tk.Label(
                category_frame,
                text=f"فئة {category}",
                font=self.fonts["text_bold"],
                bg=color,
                fg="white",
                padx=5, pady=5,
                width=15
            ).pack(fill=tk.X)

            # إحصائيات الفئة
            tk.Label(
                category_frame,
                text=f"عدد الدورات: {stats['courses_count']}",
                font=self.fonts["text"],
                bg=self.colors["light"]
            ).pack(anchor=tk.W, pady=2)

            tk.Label(
                category_frame,
                text=f"إجمالي الطلاب: {stats['total_students']}",
                font=self.fonts["text"],
                bg=self.colors["light"]
            ).pack(anchor=tk.W, pady=2)

            tk.Label(
                category_frame,
                text=f"المستبعدون: {stats['excluded_students']}",
                font=self.fonts["text"],
                bg=self.colors["light"]
            ).pack(anchor=tk.W, pady=2)

            tk.Label(
                category_frame,
                text=f"الخريجون: {stats['graduates']}",
                font=self.fonts["text"],
                bg=self.colors["light"]
            ).pack(anchor=tk.W, pady=2)

        # إنشاء الإطار الرئيسي
        main_frame = tk.Frame(archive_window, bg=self.colors["light"])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # تقسيم الإطار إلى قسمين (الدورات على اليمين، التفاصيل على اليسار)
        left_frame = tk.Frame(main_frame, bg=self.colors["light"], width=800)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        right_frame = tk.LabelFrame(main_frame, text="الدورات في الأرشيف", font=self.fonts["subtitle"],
                                    bg=self.colors["light"], fg=self.colors["dark"], width=410)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))

        # قائمة الدورات
        courses_frame = tk.Frame(right_frame, bg=self.colors["light"])
        courses_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        courses_scroll = tk.Scrollbar(courses_frame)
        courses_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        courses_listbox = tk.Listbox(
            courses_frame,
            font=self.fonts["text"],
            selectbackground=self.colors["primary"],
            selectforeground="white",
            yscrollcommand=courses_scroll.set
        )
        courses_listbox.pack(fill=tk.BOTH, expand=True)
        courses_scroll.config(command=courses_listbox.yview)

        # إضافة الدورات إلى القائمة مع فئة الدورة
        for course_name in self.archive_data["courses"].keys():
            course_students = self.archive_data["courses"][course_name]["students"]
            course_category = self.archive_data["courses"][course_name]["course_info"].get("category", "غير محدد")
            course_display = f"{course_name} ({len(course_students)} طالب) - فئة: {course_category}"
            courses_listbox.insert(tk.END, course_display)

        # تبويبات عرض بيانات الدورة المحددة
        details_notebook = ttk.Notebook(left_frame)
        details_notebook.pack(fill=tk.BOTH, expand=True)

        # تبويبات التفاصيل
        students_tab = tk.Frame(details_notebook, bg=self.colors["light"])
        attendance_tab = tk.Frame(details_notebook, bg=self.colors["light"])
        stats_tab = tk.Frame(details_notebook, bg=self.colors["light"])

        details_notebook.add(students_tab, text="قائمة الطلاب")
        details_notebook.add(attendance_tab, text="سجلات الحضور")
        details_notebook.add(stats_tab, text="الإحصائيات")

        # جدول عرض الطلاب
        students_frame = tk.Frame(students_tab, bg=self.colors["light"])
        students_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        students_scroll = tk.Scrollbar(students_frame)
        students_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        students_tree = ttk.Treeview(
            students_frame,
            columns=("id", "name", "rank", "phone", "status"),
            show="headings",
            yscrollcommand=students_scroll.set,
            style="Bold.Treeview"
        )

        students_tree.column("id", width=120, anchor=tk.CENTER)
        students_tree.column("name", width=180, anchor=tk.CENTER)
        students_tree.column("rank", width=120, anchor=tk.CENTER)
        students_tree.column("phone", width=120, anchor=tk.CENTER)
        students_tree.column("status", width=80, anchor=tk.CENTER)

        students_tree.heading("id", text="رقم الهوية")
        students_tree.heading("name", text="الاسم")
        students_tree.heading("rank", text="الرتبة")
        students_tree.heading("phone", text="رقم الجوال")
        students_tree.heading("status", text="الحالة")

        students_tree.pack(fill=tk.BOTH, expand=True)
        students_scroll.config(command=students_tree.yview)

        # إضافة ألوان مميزة للطلاب حسب الحالة
        students_tree.tag_configure("excluded", background="#ffcdd2")  # لون أحمر فاتح للمستبعدين
        students_tree.tag_configure("active", background="#e8f5e9")  # لون أخضر فاتح للموجودين

        # جدول عرض سجلات الحضور
        attendance_frame = tk.Frame(attendance_tab, bg=self.colors["light"])
        attendance_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        attendance_scroll = tk.Scrollbar(attendance_frame)
        attendance_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        attendance_tree = ttk.Treeview(
            attendance_frame,
            columns=("id", "name", "date", "status", "time"),
            show="headings",
            yscrollcommand=attendance_scroll.set,
            style="Bold.Treeview"
        )

        attendance_tree.column("id", width=120, anchor=tk.CENTER)
        attendance_tree.column("name", width=180, anchor=tk.CENTER)
        attendance_tree.column("date", width=100, anchor=tk.CENTER)
        attendance_tree.column("status", width=100, anchor=tk.CENTER)
        attendance_tree.column("time", width=100, anchor=tk.CENTER)

        attendance_tree.heading("id", text="رقم الهوية")
        attendance_tree.heading("name", text="الاسم")
        attendance_tree.heading("date", text="التاريخ")
        attendance_tree.heading("status", text="الحالة")
        attendance_tree.heading("time", text="الوقت")

        attendance_tree.pack(fill=tk.BOTH, expand=True)
        attendance_scroll.config(command=attendance_tree.yview)

        # إعداد تبويب الإحصائيات
        stats_title_frame = tk.Frame(stats_tab, bg=self.colors["light"])
        stats_title_frame.pack(fill=tk.X, pady=(10, 5))

        stats_title_label = tk.Label(
            stats_title_frame,
            text="اختر دورة لعرض الإحصائيات",
            font=self.fonts["title"],
            bg=self.colors["light"]
        )
        stats_title_label.pack()

        # إضافة معلومات فئة الدورة
        category_label = tk.Label(
            stats_title_frame,
            text="فئة الدورة: غير محدد",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["primary"]
        )
        category_label.pack(pady=5)

        # إطار إحصائيات الحضور والغياب
        attendance_stats_frame = tk.LabelFrame(
            stats_tab,
            text="إحصائيات الحضور والغياب",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["dark"]
        )
        attendance_stats_frame.pack(fill=tk.X, padx=10, pady=5)

        attendance_stats_inner = tk.Frame(attendance_stats_frame, bg=self.colors["light"])
        attendance_stats_inner.pack(fill=tk.X, pady=10, padx=5)

        # إطار إحصائيات الطلاب والمستبعدين
        students_stats_frame = tk.LabelFrame(
            stats_tab,
            text="إحصائيات الطلاب والخريجين",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["dark"]
        )
        students_stats_frame.pack(fill=tk.X, padx=10, pady=5)

        students_stats_inner = tk.Frame(students_stats_frame, bg=self.colors["light"])
        students_stats_inner.pack(fill=tk.X, pady=10, padx=5)

        # دالة لعرض بيانات الدورة المحددة
        def show_course_data(event=None):
            # مسح البيانات السابقة
            for item in students_tree.get_children():
                students_tree.delete(item)

            for item in attendance_tree.get_children():
                attendance_tree.delete(item)

            # مسح الإطارات الإحصائية
            for widget in attendance_stats_inner.winfo_children():
                widget.destroy()

            for widget in students_stats_inner.winfo_children():
                widget.destroy()

            # الحصول على الدورة المحددة
            if not courses_listbox.curselection():
                return

            selected_course_text = courses_listbox.get(courses_listbox.curselection()[0])
            # استخراج اسم الدورة من النص المحدد (الاسم المعروض يتضمن عدد الطلاب والفئة)
            course_name = selected_course_text.split(" (")[0]

            if course_name not in self.archive_data["courses"]:
                return

            course_data = self.archive_data["courses"][course_name]

            # عرض فئة الدورة في تبويب الإحصائيات
            course_category = course_data["course_info"].get("category", "غير محدد")
            category_label.config(text=f"فئة الدورة: {course_category}")

            # عرض الطلاب
            excluded_count = 0  # عداد للطلاب المستبعدين

            for i, student in enumerate(course_data["students"]):
                is_excluded = student.get("is_excluded", 0)
                if is_excluded == 1:
                    excluded_count += 1
                    status = "مستبعد"
                    tag = "excluded"
                else:
                    status = "موجود"
                    tag = "active"

                item_id = students_tree.insert("", tk.END, values=(
                    student["national_id"],
                    student["name"],
                    student["rank"],
                    student["phone"],
                    status
                ))

                # تطبيق النمط المناسب حسب الحالة
                students_tree.item(item_id, tags=(tag,))

            # عرض سجلات الحضور
            for record in course_data["attendance"]:
                attendance_tree.insert("", tk.END, values=(
                    record["national_id"],
                    record["name"],
                    record["date"],
                    record["status"],
                    record["time"]
                ))

            # حساب وعرض الإحصائيات
            total_students = len(course_data["students"])
            graduates_count = total_students - excluded_count  # عدد الخريجين الفعلي

            # تجميع سجلات الحضور
            attendance_by_status = {}

            for record in course_data["attendance"]:
                status = record["status"]
                if status not in attendance_by_status:
                    attendance_by_status[status] = 0
                attendance_by_status[status] += 1

            # عرض عنوان الإحصائيات
            stats_title_label.config(text=f"إحصائيات دورة: {course_name}")

            # 1. إحصائيات الطلاب والخريجين
            student_stat_labels = [
                ("إجمالي عدد الملتحقين", total_students, self.colors["primary"]),
                ("عدد المستبعدين", excluded_count, "#D32F2F"),  # أحمر داكن للمستبعدين
                ("العدد الفعلي للخريجين", graduates_count, "#2E7D32")  # أخضر داكن للخريجين
            ]

            for title, count, color in student_stat_labels:
                stat_frame = tk.Frame(students_stats_inner, bg=self.colors["light"], bd=1, relief=tk.RIDGE, padx=5,
                                      pady=5)
                stat_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

                tk.Label(
                    stat_frame,
                    text=title,
                    font=self.fonts["text_bold"],
                    bg=color,
                    fg="white",
                    padx=5, pady=5
                ).pack(fill=tk.X)

                tk.Label(
                    stat_frame,
                    text=str(count),
                    font=self.fonts["title"],
                    bg=self.colors["light"]
                ).pack(fill=tk.X, pady=5)

            # 2. إحصائيات الحضور والغياب
            attendance_stat_labels = [
                ("عدد الحاضرين", attendance_by_status.get("حاضر", 0), self.colors["success"]),
                ("عدد الغائبين", attendance_by_status.get("غائب", 0), self.colors["danger"]),
                ("عدد المتأخرين", attendance_by_status.get("متأخر", 0), self.colors["late"]),
                ("غياب بعذر", attendance_by_status.get("غائب بعذر", 0), self.colors["excused"])
            ]

            for title, count, color in attendance_stat_labels:
                stat_frame = tk.Frame(attendance_stats_inner, bg=self.colors["light"], bd=1, relief=tk.RIDGE, padx=5,
                                      pady=5)
                stat_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

                tk.Label(
                    stat_frame,
                    text=title,
                    font=self.fonts["text_bold"],
                    bg=color,
                    fg="white",
                    padx=5, pady=5
                ).pack(fill=tk.X)

                tk.Label(
                    stat_frame,
                    text=str(count),
                    font=self.fonts["title"],
                    bg=self.colors["light"]
                ).pack(fill=tk.X, pady=5)

        # ربط وظيفة عرض البيانات بحدث اختيار الدورة
        courses_listbox.bind("<<ListboxSelect>>", show_course_data)

        # أزرار التصدير
        export_frame = tk.Frame(archive_window, bg=self.colors["light"], pady=10)
        export_frame.pack(fill=tk.X, padx=10)

        def export_course_data():
            if not courses_listbox.curselection():
                messagebox.showinfo("تنبيه", "الرجاء اختيار دورة للتصدير")
                return

            selected_course_text = courses_listbox.get(courses_listbox.curselection()[0])
            course_name = selected_course_text.split(" (")[0]

            if course_name not in self.archive_data["courses"]:
                return

            export_type = messagebox.askquestion(
                "تصدير البيانات",
                "هل تريد تصدير البيانات إلى:\n\n- نعم: ملف Excel\n- لا: ملف Word"
            )

            if export_type == "yes":
                # تصدير إلى Excel
                export_file = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    initialfile=f"بيانات_دورة_{course_name}"
                )

                if export_file:
                    try:
                        course_data = self.archive_data["courses"][course_name]

                        # تحويل بيانات الطلاب إلى DataFrame
                        students_df = pd.DataFrame(course_data["students"])

                        # تحويل بيانات الحضور إلى DataFrame
                        attendance_df = pd.DataFrame(course_data["attendance"])

                        # كتابة البيانات إلى ملف Excel
                        with pd.ExcelWriter(export_file) as writer:
                            students_df.to_excel(writer, sheet_name="الطلاب", index=False)
                            attendance_df.to_excel(writer, sheet_name="سجلات الحضور", index=False)

                            # إذا كانت هناك بيانات فصول، نصدرها أيضًا
                            if course_data["sections"]:
                                sections_df = pd.DataFrame(course_data["sections"])
                                sections_df.to_excel(writer, sheet_name="الفصول", index=False)

                            if course_data["student_sections"]:
                                student_sections_df = pd.DataFrame(course_data["student_sections"])
                                student_sections_df.to_excel(writer, sheet_name="توزيع الطلاب", index=False)

                        messagebox.showinfo("نجاح", f"تم تصدير بيانات الدورة '{course_name}' إلى Excel بنجاح")

                    except Exception as e:
                        messagebox.showerror("خطأ", f"حدث خطأ أثناء التصدير: {str(e)}")
            else:
                # تصدير إلى Word
                # هذا الخيار يحتاج إلى كتابة دالة منفصلة لإنشاء ملف Word
                messagebox.showinfo("ملاحظة", "خيار التصدير إلى Word قيد التطوير")

        export_btn = tk.Button(
            export_frame,
            text="تصدير بيانات الدورة المحددة",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=export_course_data
        )
        export_btn.pack(side=tk.LEFT, padx=5)

        # إضافة زر لتصدير الإحصائيات حسب الفئات
        export_stats_btn = tk.Button(
            export_frame,
            text="تصدير إحصائيات الفئات",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.export_category_statistics()
        )
        export_stats_btn.pack(side=tk.LEFT, padx=5)

        close_btn = tk.Button(
            export_frame,
            text="إغلاق الأرشيف",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=archive_window.destroy
        )
        close_btn.pack(side=tk.RIGHT, padx=5)

    def calculate_archive_category_stats(self):
        """حساب الإحصائيات حسب فئات الدورات في الأرشيف"""
        # القاموس الذي سيحتوي على إحصائيات كل فئة
        category_stats = {}

        # فئات الدورات المحددة مسبقاً
        categories = ["ضباط", "أفراد", "مشتركة", "مدنيين", "طلبة"]

        # تهيئة قاموس الإحصائيات لكل فئة
        for category in categories:
            category_stats[category] = {
                "courses_count": 0,  # عدد الدورات
                "total_students": 0,  # إجمالي عدد الطلاب
                "excluded_students": 0,  # عدد الطلاب المستبعدين
                "graduates": 0,  # عدد الخريجين
                "courses": []  # قائمة بأسماء الدورات
            }

        # تجميع البيانات من الأرشيف
        for course_name, course_data in self.archive_data["courses"].items():
            # الحصول على فئة الدورة، إذا كانت غير موجودة نستخدم "غير محدد"
            course_category = course_data["course_info"].get("category", "غير محدد")

            # إذا كانت الفئة غير موجودة في قائمة الفئات المحددة، نضيفها
            if course_category not in category_stats:
                category_stats[course_category] = {
                    "courses_count": 0,
                    "total_students": 0,
                    "excluded_students": 0,
                    "graduates": 0,
                    "courses": []
                }

            # زيادة عدد الدورات لهذه الفئة
            category_stats[course_category]["courses_count"] += 1

            # إضافة اسم الدورة إلى قائمة دورات الفئة
            category_stats[course_category]["courses"].append(course_name)

            # حساب عدد الطلاب والمستبعدين
            students = course_data["students"]
            total_students = len(students)
            excluded_students = sum(1 for student in students if student.get("is_excluded", 0) == 1)
            graduates = total_students - excluded_students

            # إضافة الإحصائيات إلى إحصائيات الفئة
            category_stats[course_category]["total_students"] += total_students
            category_stats[course_category]["excluded_students"] += excluded_students
            category_stats[course_category]["graduates"] += graduates

        # إزالة الفئات التي ليس لها دورات
        categories_to_remove = []
        for category in category_stats:
            if category_stats[category]["courses_count"] == 0:
                categories_to_remove.append(category)

        for category in categories_to_remove:
            del category_stats[category]

        return category_stats

    def export_category_statistics(self):
        """تصدير إحصائيات الفئات إلى ملف Excel"""
        if not hasattr(self, 'archive_data') or not self.archive_data:
            messagebox.showinfo("تنبيه", "الرجاء فتح ملف أرشيف أولاً")
            return

        # حساب الإحصائيات حسب الفئات
        category_stats = self.calculate_archive_category_stats()

        if not category_stats:
            messagebox.showinfo("تنبيه", "لا توجد إحصائيات متاحة")
            return

        # اختيار مسار الملف للتصدير
        export_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"إحصائيات_الفئات_{datetime.datetime.now().strftime('%Y%m%d')}"
        )

        if not export_file:
            return

        try:
            # إنشاء DataFrame للإحصائيات العامة
            stats_data = []
            for category, stats in category_stats.items():
                stats_data.append({
                    "الفئة": category,
                    "عدد الدورات": stats["courses_count"],
                    "إجمالي الطلاب": stats["total_students"],
                    "المستبعدون": stats["excluded_students"],
                    "الخريجون": stats["graduates"],
                    "نسبة الخريجين": f"{(stats['graduates'] / stats['total_students'] * 100):.2f}%" if stats[
                                                                                                           "total_students"] > 0 else "0%"
                })

            stats_df = pd.DataFrame(stats_data)

            # إنشاء DataFrame لتفاصيل الدورات
            courses_data = []
            for category, stats in category_stats.items():
                for course_name in stats["courses"]:
                    course_data = self.archive_data["courses"][course_name]
                    students = course_data["students"]
                    total_students = len(students)
                    excluded_students = sum(1 for student in students if student.get("is_excluded", 0) == 1)
                    graduates = total_students - excluded_students

                    courses_data.append({
                        "الفئة": category,
                        "اسم الدورة": course_name,
                        "إجمالي الطلاب": total_students,
                        "المستبعدون": excluded_students,
                        "الخريجون": graduates,
                        "نسبة الخريجين": f"{(graduates / total_students * 100):.2f}%" if total_students > 0 else "0%"
                    })

            courses_df = pd.DataFrame(courses_data)

            # تصدير البيانات إلى ملف Excel
            with pd.ExcelWriter(export_file) as writer:
                stats_df.to_excel(writer, sheet_name="الإحصائيات العامة", index=False)
                courses_df.to_excel(writer, sheet_name="تفاصيل الدورات", index=False)

                # إنشاء ورقة لكل فئة تحتوي على تفاصيل الدورات والطلاب
                for category in category_stats:
                    category_courses = category_stats[category]["courses"]

                    if not category_courses:
                        continue

                    # جمع بيانات طلاب هذه الفئة
                    category_students = []

                    for course_name in category_courses:
                        course_data = self.archive_data["courses"][course_name]

                        for student in course_data["students"]:
                            category_students.append({
                                "رقم الهوية": student["national_id"],
                                "الاسم": student["name"],
                                "الرتبة": student["rank"],
                                "الدورة": course_name,
                                "رقم الجوال": student.get("phone", ""),
                                "الحالة": "مستبعد" if student.get("is_excluded", 0) == 1 else "خريج"
                            })

                    # تصدير بيانات طلاب الفئة
                    if category_students:
                        students_df = pd.DataFrame(category_students)
                        sheet_name = f"طلاب {category}"
                        students_df.to_excel(writer, sheet_name=sheet_name, index=False)

            messagebox.showinfo("نجاح", f"تم تصدير إحصائيات الفئات بنجاح إلى:\n{export_file}")

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير الإحصائيات: {str(e)}")

    def merge_archives(self):
        """دمج ملفات أرشيف متعددة في ملف واحد"""
        # اختيار ملفات الأرشيف للدمج
        archive_files = filedialog.askopenfilenames(
            title="اختر ملفات الأرشيف للدمج",
            filetypes=[("ملفات أرشيف الدورات", "*.crsarch")]
        )

        if not archive_files:
            return

        # إنشاء أرشيف جديد
        merged_archive = {
            "metadata": {
                "creation_date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "courses_count": 0,
                "course_names": [],
                "archive_id": str(uuid.uuid4()),
                "merged_from": []
            },
            "courses": {}
        }

        try:
            # إنشاء نافذة تقدم العملية
            progress_window = tk.Toplevel(self.root)
            progress_window.title("دمج ملفات الأرشيف")
            progress_window.geometry("450x180")
            progress_window.configure(bg=self.colors["light"])
            progress_window.transient(self.root)
            progress_window.grab_set()

            # توسيط النافذة
            x = (progress_window.winfo_screenwidth() - 450) // 2
            y = (progress_window.winfo_screenheight() - 180) // 2
            progress_window.geometry(f"450x180+{x}+{y}")

            tk.Label(
                progress_window,
                text=f"جاري دمج {len(archive_files)} ملف أرشيف...",
                font=self.fonts["text_bold"],
                bg=self.colors["light"],
                pady=10
            ).pack()

            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(
                progress_window,
                variable=progress_var,
                maximum=100,
                length=400
            )
            progress_bar.pack(pady=10)

            status_label = tk.Label(
                progress_window,
                text="جاري تحضير البيانات...",
                font=self.fonts["text"],
                bg=self.colors["light"]
            )
            status_label.pack(pady=5)

            progress_window.update()

            # معالجة كل ملف أرشيف
            for i, archive_file in enumerate(archive_files):
                progress_var.set((i / len(archive_files)) * 80)
                status_label.config(text=f"جاري معالجة ملف الأرشيف {i + 1} من {len(archive_files)}...")
                progress_window.update()

                # إنشاء مجلد مؤقت لاستخراج الملف
                temp_dir = tempfile.mkdtemp()

                try:
                    # استخراج ملف الأرشيف
                    with zipfile.ZipFile(archive_file, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)

                    # قراءة ملف البيانات
                    archive_json_path = os.path.join(temp_dir, "archive_data.json")
                    with open(archive_json_path, 'r', encoding='utf-8') as f:
                        archive_data = json.load(f)

                    # إضافة معلومات هذا الأرشيف إلى قائمة الأرشيفات المدمجة
                    merged_archive["metadata"]["merged_from"].append({
                        "archive_id": archive_data["metadata"].get("archive_id", "غير معروف"),
                        "creation_date": archive_data["metadata"].get("creation_date", "غير معروف"),
                        "courses_count": archive_data["metadata"].get("courses_count", 0)
                    })

                    # دمج الدورات
                    for course_name, course_data in archive_data["courses"].items():
                        # في حالة وجود دورة بنفس الاسم، نضيف رقم للتمييز
                        new_course_name = course_name
                        counter = 1
                        while new_course_name in merged_archive["courses"]:
                            new_course_name = f"{course_name} ({counter})"
                            counter += 1

                        # إضافة الدورة إلى الأرشيف المدمج
                        merged_archive["courses"][new_course_name] = course_data
                        merged_archive["metadata"]["course_names"].append(new_course_name)
                        merged_archive["metadata"]["courses_count"] += 1

                finally:
                    # تنظيف المجلد المؤقت
                    shutil.rmtree(temp_dir)

            # تحديث شريط التقدم
            progress_var.set(85)
            status_label.config(text="جاري حفظ الأرشيف المدمج...")
            progress_window.update()

            # حفظ الأرشيف المدمج
            export_file = filedialog.asksaveasfilename(
                defaultextension=".crsarch",
                filetypes=[("ملفات أرشيف الدورات", "*.crsarch")],
                initialfile=f"merged_archive_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )

            if export_file:
                # إنشاء مجلد مؤقت للتصدير
                temp_dir = tempfile.mkdtemp()

                try:
                    # حفظ ملف JSON
                    archive_json = os.path.join(temp_dir, "archive_data.json")
                    with open(archive_json, 'w', encoding='utf-8') as f:
                        json.dump(merged_archive, f, ensure_ascii=False, indent=2)

                    # إنشاء ملف الأرشيف المضغوط
                    with zipfile.ZipFile(export_file, 'w', compression=zipfile.ZIP_DEFLATED) as archive_zip:
                        archive_zip.write(archive_json, arcname="archive_data.json")

                    progress_var.set(100)
                    status_label.config(text="تم دمج ملفات الأرشيف بنجاح!")
                    progress_window.update()

                    # إغلاق نافذة التقدم بعد ثانيتين
                    progress_window.after(2000, progress_window.destroy)

                    messagebox.showinfo("نجاح",
                                        f"تم دمج {len(archive_files)} ملف أرشيف بنجاح وحفظهم في:\n{export_file}")

                finally:
                    # تنظيف المجلد المؤقت
                    shutil.rmtree(temp_dir)
            else:
                progress_window.destroy()

        except Exception as e:
            try:
                progress_window.destroy()
            except:
                pass
            messagebox.showerror("خطأ", f"حدث خطأ أثناء دمج ملفات الأرشيف: {str(e)}")

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.enum.section import WD_ORIENTATION
    from docx.oxml.ns import qn
    from docx.oxml.shared import parse_xml, nsdecls
    from docx.oxml import OxmlElement
except ImportError:
    print("تحذير: مكتبة python-docx غير مثبتة. قم بتثبيتها باستخدام: pip install python-docx")


try:
    import arabic_reshaper
    from bidi.algorithm import get_display


    def fix_arabic(text):
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
except ImportError:
    def fix_arabic(text):
        return text

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    pdfmetrics.registerFont(TTFont('ArabicFont', 'Tajawal-Regular.ttf'))
except ImportError:
    print("تحذير: مكتبة reportlab غير مثبتة، وقد لا ينجح تصدير PDF.")


class LoginSystem:
    def __init__(self, root):
        self.root = root
        self.db_conn = self.connect_to_db()
        self.current_user = None
        self.create_users_table()
        self.create_permissions_table()
        self.check_admin_exists()
        self.setup_login_window()

    def connect_to_db(self):
        try:
            conn = sqlite3.connect("attendance.db")
            return conn
        except Exception as e:
            messagebox.showerror("خطأ في قاعدة البيانات", f"لا يمكن الاتصال بقاعدة البيانات: {str(e)}")
            exit(1)

    def create_users_table(self):
        try:
            with self.db_conn:
                try:
                    self.db_conn.execute("ALTER TABLE users DROP COLUMN role")
                except:
                    pass
                self.db_conn.execute("""
                    CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT UNIQUE,
                        password TEXT,
                        full_name TEXT,
                        created_date TEXT,
                        last_login TEXT,
                        is_active INTEGER DEFAULT 1
                    )
                """)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذّر إنشاء/تعديل جدول المستخدمين: {str(e)}")

    def create_permissions_table(self):
        try:
            with self.db_conn:
                self.db_conn.execute("""
                    CREATE TABLE IF NOT EXISTS user_permissions (
                        user_id INTEGER PRIMARY KEY,
                        can_edit_attendance INTEGER DEFAULT 1,
                        can_add_students INTEGER DEFAULT 1,
                        can_edit_students INTEGER DEFAULT 1,
                        can_delete_students INTEGER DEFAULT 0,
                        can_view_edit_history INTEGER DEFAULT 0,
                        can_reset_attendance INTEGER DEFAULT 0,
                        can_export_data INTEGER DEFAULT 1,
                        can_import_data INTEGER DEFAULT 0,
                        is_admin INTEGER DEFAULT 0,
                        FOREIGN KEY (user_id) REFERENCES users(id)
                    )
                """)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذّر إنشاء جدول الصلاحيات: {str(e)}")

    def check_admin_exists(self):
        cursor = self.db_conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users WHERE username='admin'")
        count = cursor.fetchone()[0]
        if count == 0:
            hashed_pwd = hashlib.sha256("admin123".encode()).hexdigest()
            try:
                with self.db_conn:
                    self.db_conn.execute("""
                        INSERT INTO users (username, password, full_name, created_date, is_active)
                        VALUES (?, ?, ?, ?, ?)
                    """, (
                        'admin',
                        hashed_pwd,
                        'المسؤول الرئيسي',
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        1
                    ))
                    cursor.execute("SELECT id FROM users WHERE username='admin'")
                    admin_id = cursor.fetchone()[0]
                    self.db_conn.execute("""
                        INSERT INTO user_permissions (
                            user_id, can_edit_attendance, can_add_students, 
                            can_edit_students, can_delete_students, can_view_edit_history,
                            can_reset_attendance, can_export_data, can_import_data, is_admin
                        ) VALUES (?, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                    """, (admin_id,))
            except Exception as e:
                messagebox.showerror("خطأ", f"تعذّر إنشاء حساب المدير الرئيسي: {str(e)}")

    def setup_login_window(self):
        self.colors = {
            "primary": "#1E40AF",
            "secondary": "#3B82F6",
            "background": "#F1F5F9",
            "card": "#FFFFFF",
            "text": "#1F2937",
            "border": "#E5E7EB",
            "error": "#EF4444"
        }
        self.fonts = {
            "heading": ("Tajawal", 28, "bold"),
            "title": ("Tajawal", 18, "bold"),
            "normal": ("Tajawal", 14),
            "bold": ("Tajawal", 14, "bold"),
            "small": ("Tajawal", 12)
        }

        self.root.title("نظام الحضور والغياب - تسجيل الدخول")
        self.root.geometry("900x600")
        self.root.resizable(False, False)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 900) // 2
        y = (screen_height - 600) // 2
        self.root.geometry(f"900x600+{x}+{y}")

        main_frame = tk.Frame(self.root, bg=self.colors["background"])
        main_frame.pack(fill=tk.BOTH, expand=True)

        left_frame = tk.Frame(main_frame, bg=self.colors["primary"], width=350)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        right_frame = tk.Frame(main_frame, bg=self.colors["background"])
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        left_title = tk.Label(
            left_frame,
            text="قــســم\nشــؤون\nالمـدربـين",
            font=self.fonts["heading"],
            bg=self.colors["primary"],
            fg="white",
            justify=tk.LEFT
        )
        left_title.place(x=30, y=150)

        left_footer = tk.Label(
            left_frame,
            text="© 2025\nجميع الحقوق محفوظة \n للمهندس / عبدالرحمن جفال الشمري ",
            font=self.fonts["small"],
            bg=self.colors["primary"],
            fg="white"
        )
        left_footer.place(x=30, y=520)

        card = tk.Frame(right_frame, bg=self.colors["card"], bd=1, relief=tk.RIDGE, padx=40, pady=30)
        card.place(relx=0.5, rely=0.5, anchor=tk.CENTER, width=420, height=380)

        login_label = tk.Label(card, text="تسجيل الدخول", font=self.fonts["title"], fg=self.colors["primary"],
                               bg=self.colors["card"])
        login_label.pack(pady=(0, 20))

        username_label = tk.Label(card, text="اسم المستخدم:", font=self.fonts["bold"], bg=self.colors["card"],
                                  fg=self.colors["text"])
        username_label.pack(anchor="w", pady=(5, 0))

        self.username_entry = tk.Entry(card, font=self.fonts["normal"], bg=self.colors["card"], fg=self.colors["text"],
                                       highlightthickness=1, highlightbackground=self.colors["border"], relief=tk.FLAT)
        self.username_entry.pack(fill=tk.X, pady=(0, 10), ipady=6)
        self.username_entry.focus_set()

        password_label = tk.Label(card, text="كلمة المرور:", font=self.fonts["bold"], bg=self.colors["card"],
                                  fg=self.colors["text"])
        password_label.pack(anchor="w", pady=(5, 0))

        self.password_entry = tk.Entry(card, font=self.fonts["normal"], bg=self.colors["card"], fg=self.colors["text"],
                                       highlightthickness=1, highlightbackground=self.colors["border"], show="•",
                                       relief=tk.FLAT)
        self.password_entry.pack(fill=tk.X, pady=(0, 20), ipady=6)
        self.password_entry.bind("<Return>", lambda event: self.login())

        login_button = tk.Button(card, text="دخول", font=self.fonts["bold"], bg=self.colors["secondary"], fg="white",
                                 bd=0, relief=tk.FLAT, cursor="hand2", command=self.login)
        login_button.pack(fill=tk.X, pady=(0, 10), ipady=8)

        self.status_label = tk.Label(card, text="", font=self.fonts["small"], bg=self.colors["card"],
                                     fg=self.colors["error"])
        self.status_label.pack()

    def login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            messagebox.showwarning("تنبيه", "الرجاء إدخال اسم المستخدم وكلمة المرور")
            return

        hashed_pwd = hashlib.sha256(password.encode()).hexdigest()
        cursor = self.db_conn.cursor()
        cursor.execute("""
            SELECT u.id, u.username, u.full_name
            FROM users u
            WHERE u.username=? AND u.password=? AND u.is_active=1
        """, (username, hashed_pwd))
        user = cursor.fetchone()

        if user:
            cursor.execute("""
                SELECT * FROM user_permissions WHERE user_id=?
            """, (user[0],))
            permissions = cursor.fetchone()

            if not permissions:
                is_admin = 1 if username == 'admin' else 0
                with self.db_conn:
                    cursor.execute("""
                        INSERT INTO user_permissions (
                            user_id, can_edit_attendance, can_add_students, 
                            can_edit_students, can_delete_students, can_view_edit_history,
                            can_reset_attendance, can_export_data, can_import_data, is_admin
                        ) VALUES (?, 1, 1, 1, ?, ?, ?, 1, ?, ?)
                    """, (user[0], is_admin, is_admin, is_admin, is_admin, is_admin))

                cursor.execute("SELECT * FROM user_permissions WHERE user_id=?", (user[0],))
                permissions = cursor.fetchone()

            with self.db_conn:
                self.db_conn.execute("UPDATE users SET last_login=? WHERE id=?",
                                     (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user[0]))

            self.current_user = {
                "id": user[0],
                "username": user[1],
                "full_name": user[2],
                "permissions": {
                    "can_edit_attendance": bool(permissions[1]),
                    "can_add_students": bool(permissions[2]),
                    "can_edit_students": bool(permissions[3]),
                    "can_delete_students": bool(permissions[4]),
                    "can_view_edit_history": bool(permissions[5]),
                    "can_reset_attendance": bool(permissions[6]),
                    "can_export_data": bool(permissions[7]),
                    "can_import_data": bool(permissions[8]),
                    "is_admin": bool(permissions[9])
                }
            }

            self.root.destroy()

            new_root = tk.Tk()
            ModernAttendanceSystem(new_root, self.current_user, self.db_conn)
            new_root.mainloop()
        else:
            messagebox.showwarning("خطأ", "اسم المستخدم أو كلمة المرور غير صحيحة")


class UserManagement:
    def __init__(self, root, conn, current_user, colors, fonts):
        self.root = root
        self.conn = conn
        self.current_user = current_user
        self.colors = colors
        self.fonts = fonts
        self.create_user_management_window()

    def create_user_management_window(self):
        self.user_window = tk.Toplevel(self.root)
        self.user_window.title("إدارة المستخدمين")
        self.user_window.geometry("900x700")
        self.user_window.configure(bg=self.colors["light"])
        # self.user_window.transient(self.root)  # قم بتعليق هذا السطر أو حذفه
        self.user_window.grab_set()

        # تفعيل خاصية تغيير حجم النافذة
        self.user_window.resizable(True, True)

        x = (self.user_window.winfo_screenwidth() - 900) // 2
        y = (self.user_window.winfo_screenheight() - 700) // 2
        self.user_window.geometry(f"900x700+{x}+{y}")

        tk.Label(
            self.user_window,
            text="إدارة المستخدمين",
            font=self.fonts["large_title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10, width=900
        ).pack(fill=tk.X)

        button_frame = tk.Frame(self.user_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=10)

        add_user_btn = tk.Button(
            button_frame, text="إضافة مستخدم جديد", font=self.fonts["text_bold"], bg=self.colors["success"], fg="white",
            padx=10, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=self.add_user
        )
        add_user_btn.pack(side=tk.RIGHT, padx=5)

        edit_user_btn = tk.Button(
            button_frame, text="تعديل المستخدم المحدد", font=self.fonts["text_bold"], bg=self.colors["warning"],
            fg="white",
            padx=10, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=self.edit_user
        )
        edit_user_btn.pack(side=tk.RIGHT, padx=5)

        toggle_active_btn = tk.Button(
            button_frame, text="تفعيل/تعطيل المستخدم", font=self.fonts["text_bold"], bg=self.colors["secondary"],
            fg="white",
            padx=10, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=self.toggle_user_active
        )
        toggle_active_btn.pack(side=tk.RIGHT, padx=5)

        delete_user_btn = tk.Button(
            button_frame, text="حذف المستخدم المحدد", font=self.fonts["text_bold"], bg=self.colors["danger"],
            fg="white",
            padx=10, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=self.delete_user
        )
        delete_user_btn.pack(side=tk.RIGHT, padx=5)

        manage_permissions_btn = tk.Button(
            button_frame, text="إدارة صلاحيات المستخدم", font=self.fonts["text_bold"], bg="#9333EA", fg="white",
            padx=10, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=self.manage_user_permissions
        )
        manage_permissions_btn.pack(side=tk.RIGHT, padx=5)

        table_frame = tk.Frame(self.user_window, bg=self.colors["light"], padx=10, pady=10)
        table_frame.pack(fill=tk.BOTH, expand=True)

        tree_scroll = tk.Scrollbar(table_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.users_tree = ttk.Treeview(
            table_frame,
            columns=("id", "username", "full_name", "created_date", "last_login", "status", "is_admin"),
            show="headings",
            yscrollcommand=tree_scroll.set
        )
        self.users_tree.column("id", width=50, anchor=tk.CENTER)
        self.users_tree.column("username", width=120, anchor=tk.CENTER)
        self.users_tree.column("full_name", width=150, anchor=tk.CENTER)
        self.users_tree.column("created_date", width=120, anchor=tk.CENTER)
        self.users_tree.column("last_login", width=120, anchor=tk.CENTER)
        self.users_tree.column("status", width=80, anchor=tk.CENTER)
        self.users_tree.column("is_admin", width=80, anchor=tk.CENTER)

        self.users_tree.heading("id", text="الرقم")
        self.users_tree.heading("username", text="اسم المستخدم")
        self.users_tree.heading("full_name", text="الاسم الكامل")
        self.users_tree.heading("created_date", text="تاريخ الإنشاء")
        self.users_tree.heading("last_login", text="آخر تسجيل دخول")
        self.users_tree.heading("status", text="الحالة")
        self.users_tree.heading("is_admin", text="مشرف")

        self.users_tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.users_tree.yview)

        self.users_tree.tag_configure("active", background="#e8f5e9")
        self.users_tree.tag_configure("inactive", background="#ffebee")
        self.users_tree.tag_configure("admin", background="#e1f5fe")

        self.load_users()

        close_btn = tk.Button(
            self.user_window, text="إغلاق", font=self.fonts["text_bold"], bg=self.colors["dark"], fg="white",
            padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=self.user_window.destroy
        )
        close_btn.pack(pady=10)

    def load_users(self):
        for item in self.users_tree.get_children():
            self.users_tree.delete(item)
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT u.id, u.username, u.full_name, u.created_date, u.last_login, u.is_active,
                   COALESCE(p.is_admin, 0) as is_admin
            FROM users u
            LEFT JOIN user_permissions p ON u.id = p.user_id
        """)
        users = cursor.fetchall()
        for user in users:
            user_id, username, full_name, created_date, last_login, is_active, is_admin = user
            status = "نشط" if is_active else "معطل"
            admin_status = "نعم" if is_admin else "لا"
            if not last_login:
                last_login = "لم يسجل الدخول بعد"
            item_id = self.users_tree.insert("", tk.END, values=(
                user_id, username, full_name, created_date, last_login, status, admin_status))

            if not is_active:
                self.users_tree.item(item_id, tags=("inactive",))
            elif is_admin:
                self.users_tree.item(item_id, tags=("admin",))
            else:
                self.users_tree.item(item_id, tags=("active",))

    def add_user(self):
        add_window = tk.Toplevel(self.user_window)
        add_window.title("إضافة مستخدم جديد")
        add_window.geometry("400x430")
        add_window.configure(bg=self.colors["light"])
        add_window.transient(self.user_window)
        add_window.grab_set()

        x = (add_window.winfo_screenwidth() - 400) // 2
        y = (add_window.winfo_screenheight() - 430) // 2
        add_window.geometry(f"400x430+{x}+{y}")

        tk.Label(
            add_window,
            text="إضافة مستخدم جديد",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10, width=400
        ).pack(fill=tk.X)

        form_frame = tk.Frame(add_window, bg=self.colors["light"], padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH)

        tk.Label(form_frame, text="اسم المستخدم:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=0, column=1, padx=5, pady=8, sticky=tk.E)
        username_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        username_entry.grid(row=0, column=0, padx=5, pady=8, sticky=tk.W)

        tk.Label(form_frame, text="الاسم الكامل:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=1, column=1, padx=5, pady=8, sticky=tk.E)
        fullname_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        fullname_entry.grid(row=1, column=0, padx=5, pady=8, sticky=tk.W)

        tk.Label(form_frame, text="كلمة المرور:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=2, column=1, padx=5, pady=8, sticky=tk.E)
        password_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25, show="*")
        password_entry.grid(row=2, column=0, padx=5, pady=8, sticky=tk.W)

        tk.Label(form_frame, text="تأكيد كلمة المرور:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=3, column=1, padx=5, pady=8, sticky=tk.E)
        confirm_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25, show="*")
        confirm_entry.grid(row=3, column=0, padx=5, pady=8, sticky=tk.W)

        is_admin_var = tk.IntVar(value=0)
        admin_check = tk.Checkbutton(
            form_frame,
            text="جعل هذا المستخدم مشرفًا",
            variable=is_admin_var,
            font=self.fonts["text"],
            bg=self.colors["light"]
        )
        admin_check.grid(row=4, column=0, columnspan=2, padx=5, pady=8, sticky=tk.W)

        button_frame = tk.Frame(add_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X)

        def save_user():
            username = username_entry.get().strip()
            fullname = fullname_entry.get().strip()
            password = password_entry.get().strip()
            confirm = confirm_entry.get().strip()
            is_admin = is_admin_var.get()

            if not all([username, fullname, password, confirm]):
                messagebox.showwarning("تنبيه", "يجب ملء جميع الحقول")
                return
            if password != confirm:
                messagebox.showwarning("تنبيه", "كلمات المرور غير متطابقة")
                return
            if not re.match(r'^[a-zA-Z0-9_-]+$', username):
                messagebox.showwarning("تنبيه", "اسم المستخدم يجب أن يتكون من حروف إنجليزية وأرقام وشرطات فقط")
                return
            cursor = self.conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM users WHERE username=?", (username,))
            count = cursor.fetchone()[0]
            if count > 0:
                messagebox.showwarning("تنبيه", "اسم المستخدم موجود بالفعل")
                return

            hashed_pwd = hashlib.sha256(password.encode()).hexdigest()
            try:
                with self.conn:
                    self.conn.execute("""
                        INSERT INTO users (username, password, full_name, created_date, is_active)
                        VALUES (?, ?, ?, ?, ?)
                    """, (
                        username,
                        hashed_pwd,
                        fullname,
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        1
                    ))

                    cursor.execute("SELECT id FROM users WHERE username=?", (username,))
                    user_id = cursor.fetchone()[0]

                    if is_admin:
                        self.conn.execute("""
                            INSERT INTO user_permissions (
                                user_id, can_edit_attendance, can_add_students, 
                                can_edit_students, can_delete_students, can_view_edit_history,
                                can_reset_attendance, can_export_data, can_import_data, is_admin
                            ) VALUES (?, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                        """, (user_id,))
                    else:
                        self.conn.execute("""
                            INSERT INTO user_permissions (
                                user_id, can_edit_attendance, can_add_students, 
                                can_edit_students, can_delete_students, can_view_edit_history,
                                can_reset_attendance, can_export_data, can_import_data, is_admin
                            ) VALUES (?, 1, 1, 1, 0, 0, 0, 1, 0, 0)
                        """, (user_id,))

                messagebox.showinfo("نجاح", "تم إضافة المستخدم بنجاح")
                add_window.destroy()
                self.load_users()
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء إضافة المستخدم: {str(e)}")

        save_btn = tk.Button(button_frame, text="حفظ", font=self.fonts["text_bold"], bg=self.colors["success"],
                             fg="white",
                             padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=save_user)
        save_btn.pack(side=tk.LEFT, padx=10)
        cancel_btn = tk.Button(button_frame, text="إلغاء", font=self.fonts["text_bold"], bg=self.colors["danger"],
                               fg="white",
                               padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=add_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=10)

    def edit_user(self):
        selected_item = self.users_tree.selection()
        if not selected_item:
            messagebox.showinfo("تنبيه", "الرجاء تحديد مستخدم من القائمة")
            return
        values = self.users_tree.item(selected_item, "values")
        user_id = values[0]
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM users WHERE id=?", (user_id,))
        user = cursor.fetchone()
        if not user:
            messagebox.showerror("خطأ", "لم يتم العثور على المستخدم")
            return

        cursor.execute("SELECT * FROM user_permissions WHERE user_id=?", (user_id,))
        permissions = cursor.fetchone()
        is_admin = 0
        if permissions:
            is_admin = permissions[9]

        edit_window = tk.Toplevel(self.user_window)
        edit_window.title("تعديل المستخدم")
        edit_window.geometry("400x430")
        edit_window.configure(bg=self.colors["light"])
        edit_window.transient(self.user_window)
        edit_window.grab_set()

        x = (edit_window.winfo_screenwidth() - 400) // 2
        y = (edit_window.winfo_screenheight() - 430) // 2
        edit_window.geometry(f"400x430+{x}+{y}")

        tk.Label(
            edit_window,
            text=f"تعديل المستخدم: {user[1]}",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10, width=400
        ).pack(fill=tk.X)

        form_frame = tk.Frame(edit_window, bg=self.colors["light"], padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH)

        tk.Label(form_frame, text="اسم المستخدم:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=0, column=1, padx=5, pady=8, sticky=tk.E)
        username_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        username_entry.insert(0, user[1])
        username_entry.grid(row=0, column=0, padx=5, pady=8, sticky=tk.W)

        tk.Label(form_frame, text="الاسم الكامل:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=1, column=1, padx=5, pady=8, sticky=tk.E)
        fullname_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        fullname_entry.insert(0, user[3])
        fullname_entry.grid(row=1, column=0, padx=5, pady=8, sticky=tk.W)

        tk.Label(form_frame, text="كلمة المرور الجديدة:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=2, column=1, padx=5, pady=8, sticky=tk.E)
        password_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25, show="*")
        password_entry.grid(row=2, column=0, padx=5, pady=8, sticky=tk.W)

        tk.Label(form_frame, text="تأكيد كلمة المرور:", font=self.fonts["text_bold"], bg=self.colors["light"],
                 anchor=tk.E).grid(row=3, column=1, padx=5, pady=8, sticky=tk.E)
        confirm_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25, show="*")
        confirm_entry.grid(row=3, column=0, padx=5, pady=8, sticky=tk.W)

        is_admin_var = tk.IntVar(value=is_admin)
        admin_check = tk.Checkbutton(
            form_frame,
            text="هذا المستخدم مشرف",
            variable=is_admin_var,
            font=self.fonts["text"],
            bg=self.colors["light"]
        )
        admin_check.grid(row=4, column=0, columnspan=2, padx=5, pady=8, sticky=tk.W)

        button_frame = tk.Frame(edit_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X)

        def save_changes():
            username = username_entry.get().strip()
            fullname = fullname_entry.get().strip()
            password = password_entry.get().strip()
            confirm = confirm_entry.get().strip()
            is_admin = is_admin_var.get()

            if not all([username, fullname]):
                messagebox.showwarning("تنبيه", "يجب ملء الحقول الأساسية")
                return

            if password:
                if password != confirm:
                    messagebox.showwarning("تنبيه", "كلمات المرور غير متطابقة")
                    return
            try:
                with self.conn:
                    if password:
                        hashed_pwd = hashlib.sha256(password.encode()).hexdigest()
                        self.conn.execute("UPDATE users SET username=?, full_name=?, password=? WHERE id=?",
                                          (username, fullname, hashed_pwd, user[0]))
                    else:
                        self.conn.execute("UPDATE users SET username=?, full_name=? WHERE id=?",
                                          (username, fullname, user[0]))

                    cursor = self.conn.cursor()
                    cursor.execute("SELECT COUNT(*) FROM user_permissions WHERE user_id=?", (user[0],))
                    has_permissions = cursor.fetchone()[0] > 0

                    if has_permissions:
                        self.conn.execute("UPDATE user_permissions SET is_admin=? WHERE user_id=?",
                                          (is_admin, user[0]))
                    else:
                        if is_admin:
                            self.conn.execute("""
                                INSERT INTO user_permissions (
                                    user_id, can_edit_attendance, can_add_students, 
                                    can_edit_students, can_delete_students, can_view_edit_history,
                                    can_reset_attendance, can_export_data, can_import_data, is_admin
                                ) VALUES (?, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                            """, (user[0],))
                        else:
                            self.conn.execute("""
                                INSERT INTO user_permissions (
                                    user_id, can_edit_attendance, can_add_students, 
                                    can_edit_students, can_delete_students, can_view_edit_history,
                                    can_reset_attendance, can_export_data, can_import_data, is_admin
                                ) VALUES (?, 1, 1, 1, 0, 0, 0, 1, 0, 0)
                            """, (user[0],))

                messagebox.showinfo("نجاح", "تم تحديث بيانات المستخدم بنجاح")
                edit_window.destroy()
                self.load_users()
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء تحديث المستخدم: {str(e)}")

        save_btn = tk.Button(button_frame, text="حفظ التغييرات", font=self.fonts["text_bold"],
                             bg=self.colors["warning"], fg="white",
                             padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=save_changes)
        save_btn.pack(side=tk.LEFT, padx=10)
        cancel_btn = tk.Button(button_frame, text="إلغاء", font=self.fonts["text_bold"], bg=self.colors["danger"],
                               fg="white",
                               padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=edit_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=10)

    def toggle_user_active(self):
        selected_item = self.users_tree.selection()
        if not selected_item:
            messagebox.showinfo("تنبيه", "الرجاء تحديد مستخدم من القائمة")
            return
        values = self.users_tree.item(selected_item, "values")
        user_id = values[0]
        username = values[1]
        status_text = values[5]
        if username == self.current_user["username"]:
            messagebox.showwarning("تنبيه", "لا يمكن تعطيل المستخدم الحالي")
            return
        new_status = 0 if status_text == "نشط" else 1
        status_msg = "تفعيل" if new_status == 1 else "تعطيل"
        if not messagebox.askyesnocancel("تأكيد", f"هل تريد {status_msg} المستخدم {username}؟"):
            return
        try:
            with self.conn:
                self.conn.execute("UPDATE users SET is_active=? WHERE id=?", (new_status, user_id))
            messagebox.showinfo("نجاح", f"تم {status_msg} المستخدم بنجاح")
            self.load_users()
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ: {str(e)}")

    def delete_user(self):
        selected_item = self.users_tree.selection()
        if not selected_item:
            messagebox.showinfo("تنبيه", "الرجاء تحديد مستخدم من القائمة")
            return
        values = self.users_tree.item(selected_item, "values")
        user_id = values[0]
        username = values[1]
        if username == self.current_user["username"]:
            messagebox.showwarning("تنبيه", "لا يمكن حذف المستخدم الحالي")
            return
        if not messagebox.askyesnocancel("تأكيد", f"هل تريد حذف المستخدم {username}؟\nلا يمكن التراجع عن العملية!"):
            return
        try:
            with self.conn:
                self.conn.execute("DELETE FROM user_permissions WHERE user_id=?", (user_id,))
                self.conn.execute("DELETE FROM users WHERE id=?", (user_id,))
            messagebox.showinfo("نجاح", "تم حذف المستخدم بنجاح")
            self.load_users()
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف المستخدم: {str(e)}")

    def manage_user_permissions(self):
        selected_item = self.users_tree.selection()
        if not selected_item:
            messagebox.showinfo("تنبيه", "الرجاء تحديد مستخدم من القائمة")
            return
        values = self.users_tree.item(selected_item, "values")
        user_id = values[0]
        username = values[1]

        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM user_permissions WHERE user_id=?", (user_id,))
        permissions = cursor.fetchone()

        if not permissions:
            is_admin = 1 if values[6] == "نعم" else 0
            with self.conn:
                cursor.execute("""
                    INSERT INTO user_permissions (
                        user_id, can_edit_attendance, can_add_students, 
                        can_edit_students, can_delete_students, can_view_edit_history,
                        can_reset_attendance, can_export_data, can_import_data, is_admin
                    ) VALUES (?, 1, 1, 1, ?, ?, ?, 1, ?, ?)
                """, (user_id, is_admin, is_admin, is_admin, is_admin, is_admin))

            cursor.execute("SELECT * FROM user_permissions WHERE user_id=?", (user_id,))
            permissions = cursor.fetchone()

        perm_window = tk.Toplevel(self.user_window)
        perm_window.title(f"إدارة صلاحيات المستخدم: {username}")
        perm_window.geometry("500x550")
        perm_window.configure(bg=self.colors["light"])
        perm_window.transient(self.user_window)
        perm_window.grab_set()

        x = (perm_window.winfo_screenwidth() - 500) // 2
        y = (perm_window.winfo_screenheight() - 550) // 2
        perm_window.geometry(f"500x550+{x}+{y}")

        tk.Label(
            perm_window,
            text=f"صلاحيات المستخدم: {username}",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        perm_frame = tk.Frame(perm_window, bg=self.colors["light"], padx=20, pady=20)
        perm_frame.pack(fill=tk.BOTH, expand=True)

        is_admin_var = tk.IntVar(value=permissions[9])
        can_edit_attendance_var = tk.IntVar(value=permissions[1])
        can_add_students_var = tk.IntVar(value=permissions[2])
        can_edit_students_var = tk.IntVar(value=permissions[3])
        can_delete_students_var = tk.IntVar(value=permissions[4])
        can_view_edit_history_var = tk.IntVar(value=permissions[5])
        can_reset_attendance_var = tk.IntVar(value=permissions[6])
        can_export_data_var = tk.IntVar(value=permissions[7])
        can_import_data_var = tk.IntVar(value=permissions[8])

        def update_permissions():
            is_admin = is_admin_var.get()
            if is_admin:
                for var in [can_edit_attendance_var, can_add_students_var, can_edit_students_var,
                            can_delete_students_var, can_view_edit_history_var, can_reset_attendance_var,
                            can_export_data_var, can_import_data_var]:
                    var.set(1)

                for checkbox in permission_checkboxes:
                    checkbox.config(state=tk.DISABLED)
            else:
                for checkbox in permission_checkboxes:
                    checkbox.config(state=tk.NORMAL)

        admin_title = tk.Label(perm_frame, text="صلاحيات عامة:", font=self.fonts["text_bold"], bg=self.colors["light"])
        admin_title.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        admin_check = tk.Checkbutton(
            perm_frame,
            text="هذا المستخدم مشرف (يملك كل الصلاحيات)",
            variable=is_admin_var,
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            command=update_permissions
        )
        admin_check.grid(row=1, column=0, sticky=tk.W, pady=5)

        specific_title = tk.Label(perm_frame, text="صلاحيات محددة:", font=self.fonts["text_bold"],
                                  bg=self.colors["light"])
        specific_title.grid(row=2, column=0, sticky=tk.W, pady=(20, 10))

        permission_options = [
            (can_edit_attendance_var, "تعديل سجلات الحضور والغياب"),
            (can_add_students_var, "إضافة طلاب جدد"),
            (can_edit_students_var, "تعديل بيانات الطلاب"),
            (can_delete_students_var, "حذف الطلاب"),
            (can_view_edit_history_var, "عرض سجل التعديلات (من عدّل ومتى)"),
            (can_reset_attendance_var, "إعادة تعيين سجلات الحضور"),
            (can_export_data_var, "تصدير البيانات"),
            (can_import_data_var, "استيراد البيانات من Excel")
        ]

        permission_checkboxes = []
        for i, (var, text) in enumerate(permission_options):
            checkbox = tk.Checkbutton(
                perm_frame,
                text=text,
                variable=var,
                font=self.fonts["text"],
                bg=self.colors["light"]
            )
            checkbox.grid(row=i + 3, column=0, sticky=tk.W, pady=5)
            permission_checkboxes.append(checkbox)

        update_permissions()

        button_frame = tk.Frame(perm_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=20)

        def save_permissions():
            try:
                with self.conn:
                    self.conn.execute("""
                        UPDATE user_permissions SET
                            is_admin=?,
                            can_edit_attendance=?,
                            can_add_students=?,
                            can_edit_students=?,
                            can_delete_students=?,
                            can_view_edit_history=?,
                            can_reset_attendance=?,
                            can_export_data=?,
                            can_import_data=?
                        WHERE user_id=?
                    """, (
                        is_admin_var.get(),
                        can_edit_attendance_var.get(),
                        can_add_students_var.get(),
                        can_edit_students_var.get(),
                        can_delete_students_var.get(),
                        can_view_edit_history_var.get(),
                        can_reset_attendance_var.get(),
                        can_export_data_var.get(),
                        can_import_data_var.get(),
                        user_id
                    ))
                messagebox.showinfo("نجاح", "تم تحديث صلاحيات المستخدم بنجاح")
                perm_window.destroy()
                self.load_users()
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء تحديث الصلاحيات: {str(e)}")

        save_btn = tk.Button(button_frame, text="حفظ الصلاحيات", font=self.fonts["text_bold"],
                             bg=self.colors["success"],
                             fg="white", padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2",
                             command=save_permissions)
        save_btn.pack(side=tk.LEFT, padx=10)

        cancel_btn = tk.Button(button_frame, text="إلغاء", font=self.fonts["text_bold"], bg=self.colors["danger"],
                               fg="white", padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2",
                               command=perm_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=10)


class ModernAttendanceSystem:
    def __init__(self, root, current_user, conn=None):
        self.root = root
        self.current_user = current_user
        self.root.title("نظام تسجيل الحضور والغياب")

        # تخزين الحجم الأصلي للخطوط قبل التعديل
        self.original_fonts = {
            "large_title": ("Tajawal", 24, "bold"),
            "title": ("Tajawal", 18, "bold"),
            "subtitle": ("Tajawal", 16, "bold"),
            "text": ("Tajawal", 12),
            "text_bold": ("Tajawal", 12, "bold"),
            "small": ("Tajawal", 10)
        }

        # تعريف الألوان
        self.colors = {
            "primary": "#1a73e8",
            "secondary": "#4285f4",
            "success": "#34a853",
            "danger": "#ea4335",
            "warning": "#fbbc05",
            "light": "#f0f4f8",
            "dark": "#202124",
            "present": "#34a853",
            "absent": "#ea4335",
            "late": "#fbbc05",
            "excused": "#4285f4",
            "not_started": "#FFA500",
            "excluded": "#9C27B0",
            "field_application": "#909090",
            "student_day": "#A9A9A9",
            "evening_remote": "#A0A0A0",
            "death_case": "#7E57C2",
            "hospital": "#26A69A",
        }

        # تحديد التخطيط الأمثل بناءً على حجم الشاشة
        self.determine_best_layout()

        # تعريف الخطوط بعد تحديد الحجم المناسب
        self.fonts = self.original_fonts.copy()

        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        self.setup_styles()

        # ربط حدث تغيير حجم النافذة بدالة التكيف التلقائي
        self.root.bind('<Configure>', self.on_window_resize)

        self.tab_control = ttk.Notebook(self.root, style="Bold.TNotebook")
        self.tab_control.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        if conn:
            self.conn = conn
        else:
            self.conn = sqlite3.connect(project_structure.get_database_path())

        self.create_tables()

        self.today = datetime.datetime.now().strftime("%Y-%m-%d")

        # تعريف متغيرات الإحصائيات
        self.total_students_var = tk.StringVar(value="0")
        self.present_students_var = tk.StringVar(value="0")
        self.absent_students_var = tk.StringVar(value="0")
        self.late_students_var = tk.StringVar(value="0")
        self.excused_students_var = tk.StringVar(value="0")
        self.not_started_students_var = tk.StringVar(value="0")
        self.field_application_var = tk.StringVar(value="0")
        self.student_day_var = tk.StringVar(value="0")
        self.evening_remote_var = tk.StringVar(value="0")
        self.attendance_rate_var = tk.StringVar(value="0%")
        self.death_case_var = tk.StringVar(value="0")
        self.hospital_var = tk.StringVar(value="0")

        # تخزين إشارات لبطاقات الإحصائيات للتحكم فيها لاحقًا
        self.stats_cards = []

        self.create_header()

        self.attendance_tab = tk.Frame(self.tab_control, bg=self.colors["light"])
        self.tab_control.add(self.attendance_tab, text="سجل الحضور")

        self.attendance_log_tab = tk.Frame(self.tab_control, bg=self.colors["light"])
        self.tab_control.add(self.attendance_log_tab, text="استعراض الحضور")

        self.students_tab = tk.Frame(self.tab_control, bg=self.colors["light"])
        self.tab_control.add(self.students_tab, text="إدارة الطلاب")

        if self.current_user["permissions"]["is_admin"]:
            self.users_tab = tk.Frame(self.tab_control, bg=self.colors["light"])
            self.tab_control.add(self.users_tab, text="إدارة المستخدمين")
            self.setup_users_tab()

        self.setup_attendance_tab()
        self.setup_attendance_log_tab()
        self.setup_students_tab()

        self.status_bar = tk.Label(
            self.root,
            text=f"مرحبًا {self.current_user['full_name']} (مستخدم: {self.current_user['username']})",
            font=self.fonts["small"], bg=self.colors["primary"], fg="white", pady=5
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.archive_manager = ArchiveManager(self.root, self, self.colors, self.fonts)

        # إضافة تبويب الأرشيف
        self.archive_tab = tk.Frame(self.tab_control, bg=self.colors["light"])
        self.tab_control.add(self.archive_tab, text="أرشيف الدورات")
        self.setup_archive_tab()

        self.update_students_tree()
        self.update_statistics()
        self.update_attendance_display()

        # تطبيق التخطيط المناسب بعد إنشاء كل العناصر
        if self.screen_info["is_small_screen"]:
            self.apply_compact_layout()
        else:
            self.apply_expanded_layout()

    def determine_best_layout(self):
        """تحديد التخطيط الأمثل بناءً على إعدادات الشاشة"""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # حساب حجم النافذة المناسب (90% من حجم الشاشة مع حد أقصى)
        window_width = min(int(screen_width * 0.9), 1400)
        window_height = min(int(screen_height * 0.9), 800)

        # توسيط النافذة
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # تعيين حجم وموقع النافذة
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # تعيين الحد الأدنى لحجم النافذة
        self.root.minsize(800, 600)

        # حفظ معلومات الشاشة لاستخدامها لاحقاً
        self.screen_info = {
            "screen_width": screen_width,
            "screen_height": screen_height,
            "window_width": window_width,
            "window_height": window_height,
            "is_small_screen": screen_width < 1200,
            "is_high_dpi": screen_width > 2000,
            "scale_factor": min(window_width / 1366, window_height / 768)  # عامل القياس النسبي
        }

        # تعديل أحجام الخطوط بناءً على عامل القياس إذا كانت شاشة عالية الدقة
        if self.screen_info["is_high_dpi"]:
            self.adjust_font_sizes(self.screen_info["scale_factor"])

    def setup_styles(self):
        """إعداد أنماط العناصر الرسومية"""
        self.style = ttk.Style()  # ✅ ضروري تعريف الكائن قبل الاستخدام

        self.style.configure("Bold.TNotebook.Tab", font=self.fonts["subtitle"])
        self.style.configure(
            "Bold.Treeview",
            background=self.colors["light"],
            foreground=self.colors["dark"],
            rowheight=30,
            fieldbackground=self.colors["light"],
            font=self.fonts["text_bold"]
        )
        self.style.configure(
            "Bold.Treeview.Heading",
            font=self.fonts["text_bold"],
            background=self.colors["primary"],
            foreground="white"
        )
        self.style.map('Bold.Treeview', background=[('selected', self.colors["primary"])])

        self.style.configure(
            "Profile.Treeview",
            background=self.colors["light"],
            foreground=self.colors["dark"],
            rowheight=32,
            fieldbackground=self.colors["light"],
            font=self.fonts["text_bold"]
        )
        self.style.configure(
            "Profile.Treeview.Heading",
            font=self.fonts["subtitle"],
            background=self.colors["primary"],
            foreground="white"
        )

    def on_window_resize(self, event=None):
        """تستجيب لتغيير حجم النافذة وتعدل العناصر تلقائياً"""
        # تجاهل الأحداث الصغيرة جدًا لتحسين الأداء
        if hasattr(self, 'last_width') and hasattr(self, 'last_height'):
            width_diff = abs(self.root.winfo_width() - self.last_width)
            height_diff = abs(self.root.winfo_height() - self.last_height)
            if width_diff < 10 and height_diff < 10:
                return

        # تخزين الحجم الحالي
        self.last_width = self.root.winfo_width()
        self.last_height = self.root.winfo_height()

        # تحديث معلومات الشاشة
        self.screen_info["window_width"] = self.last_width
        self.screen_info["window_height"] = self.last_height
        self.screen_info["is_small_screen"] = self.last_width < 1200

        # تعديل عرض الأعمدة في الجداول
        self.adjust_column_widths()

        # تعديل حجم النصوص في علامات التبويب
        self.adjust_tab_text()

        # تطبيق التخطيط المناسب
        if self.screen_info["is_small_screen"]:
            self.apply_compact_layout()
        else:
            self.apply_expanded_layout()

    def adjust_font_sizes(self, scale_factor):
        """تعديل أحجام الخطوط بناءً على عامل القياس"""
        # تحديث قيم الخطوط بناءً على عامل القياس
        self.fonts = {
            "large_title": ("Tajawal", int(self.original_fonts["large_title"][1] * scale_factor), "bold"),
            "title": ("Tajawal", int(self.original_fonts["title"][1] * scale_factor), "bold"),
            "subtitle": ("Tajawal", int(self.original_fonts["subtitle"][1] * scale_factor), "bold"),
            "text": ("Tajawal", int(self.original_fonts["text"][1] * scale_factor)),
            "text_bold": ("Tajawal", int(self.original_fonts["text_bold"][1] * scale_factor), "bold"),
            "small": ("Tajawal", int(self.original_fonts["small"][1] * scale_factor))
        }

        # تحديث أنماط العناصر الرسومية
        self.setup_styles()

    def adjust_column_widths(self):
        """تعديل عرض الأعمدة في جداول العرض بناءً على حجم النافذة"""
        try:
            # تعديل جدول سجل الحضور
            if hasattr(self, 'attendance_tree'):
                available_width = self.attendance_tree.winfo_width()
                if available_width > 50:  # تأكد من تهيئة العنصر
                    # تحديد النسب المئوية للأعمدة - زيادة نسبة عمود الاسم
                    col_ratios = [0.12, 0.28, 0.10, 0.12, 0.10, 0.10, 0.10, 0.08]  # زيادة عرض الاسم من 0.20 إلى 0.28

                    # حساب العرض الفعلي لكل عمود
                    for i, ratio in enumerate(col_ratios):
                        width = int(available_width * ratio)
                        if width > 10:  # تجنب القيم السالبة أو الصغيرة جدًا
                            self.attendance_tree.column(self.attendance_tree["columns"][i], width=width)

            # تعديل جدول الطلاب
            if hasattr(self, 'students_tree'):
                available_width = self.students_tree.winfo_width()
                if available_width > 50:
                    col_ratios = [0.15, 0.35, 0.15, 0.15, 0.15, 0.05]  # زيادة عرض الاسم من 0.30 إلى 0.35
                    for i, ratio in enumerate(col_ratios):
                        width = int(available_width * ratio)
                        if width > 10:
                            self.students_tree.column(self.students_tree["columns"][i], width=width)
        except Exception as e:
            print(f"خطأ عند تعديل عرض الأعمدة: {str(e)}")

    def adjust_tab_text(self):
        """تعديل نصوص علامات التبويب حسب المساحة المتاحة"""
        window_width = self.root.winfo_width()

        # على الشاشات الصغيرة، استخدم أسماء مختصرة
        if window_width < 800:
            self.tab_control.tab(0, text="الحضور")
            self.tab_control.tab(1, text="السجل")
            self.tab_control.tab(2, text="الطلاب")
            if self.current_user["permissions"]["is_admin"]:
                self.tab_control.tab(3, text="المستخدمين")
                self.tab_control.tab(4, text="الأرشيف")
            else:
                self.tab_control.tab(3, text="الأرشيف")
        else:
            # على الشاشات الكبيرة، استخدم الأسماء الكاملة
            self.tab_control.tab(0, text="سجل الحضور")
            self.tab_control.tab(1, text="استعراض الحضور")
            self.tab_control.tab(2, text="إدارة الطلاب")
            if self.current_user["permissions"]["is_admin"]:
                self.tab_control.tab(3, text="إدارة المستخدمين")
                self.tab_control.tab(4, text="أرشيف الدورات")
            else:
                self.tab_control.tab(3, text="أرشيف الدورات")

    def apply_compact_layout(self):
        """تطبيق التخطيط المضغوط للشاشات الصغيرة"""
        # تخزين وضع التخطيط الحالي
        self.current_layout = "compact"

        # تنظيم الإحصائيات في عمود واحد
        self.organize_stats_in_one_column()

        # تعديل عدد الأزرار المعروضة
        self.organize_buttons_for_small_screen()

    def apply_expanded_layout(self):
        """تطبيق التخطيط الموسع للشاشات الكبيرة"""
        # تخزين وضع التخطيط الحالي
        self.current_layout = "expanded"

        # تنظيم الإحصائيات في صفين
        self.organize_stats_in_two_rows()

        # عرض كامل للأزرار
        self.show_all_buttons()

    def organize_stats_in_one_column(self):
        """تنظيم بطاقات الإحصائيات في عمود واحد للشاشات الصغيرة"""
        # التنفيذ فقط إذا كان التخطيط الحالي ليس مضغوطًا
        if hasattr(self, 'current_layout') and self.current_layout == "compact":
            return

        if hasattr(self, 'stats_cards') and self.stats_cards:
            stats_frame = self.find_parent_frame(self.stats_cards[0])

            if stats_frame:
                # إزالة الصفوف القديمة
                for child in stats_frame.winfo_children():
                    if child != self.stats_cards[0].master:  # حفظ الإطار الرئيسي
                        child.destroy()

                # إنشاء إطار واحد للعمود
                column_frame = tk.Frame(stats_frame, bg=self.colors["light"])
                column_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

                # إعادة تنظيم بطاقات الإحصائيات
                for i, card in enumerate(self.stats_cards):
                    card.pack_forget()  # إزالة من التخطيط الحالي
                    card.pack(in_=column_frame, fill=tk.X, padx=5, pady=2)  # إعادة تنظيم في العمود

    def organize_stats_in_two_rows(self):
        """تنظيم بطاقات الإحصائيات في صفين للشاشات الكبيرة"""
        # التنفيذ فقط إذا كان التخطيط الحالي ليس موسعًا
        if hasattr(self, 'current_layout') and self.current_layout == "expanded":
            return

        if hasattr(self, 'stats_cards') and self.stats_cards:
            stats_frame = self.find_parent_frame(self.stats_cards[0])

            if stats_frame:
                # إزالة العمود القديم
                for child in stats_frame.winfo_children():
                    child.destroy()

                # إنشاء إطارين للصفين
                top_counter_frame = tk.Frame(stats_frame, bg=self.colors["light"])
                top_counter_frame.pack(fill=tk.X, padx=5, pady=5)

                bottom_counter_frame = tk.Frame(stats_frame, bg=self.colors["light"])
                bottom_counter_frame.pack(fill=tk.X, padx=5, pady=5)

                # توزيع بطاقات الإحصائيات على الصفين
                half_count = len(self.stats_cards) // 2

                for i, card in enumerate(self.stats_cards):
                    card.pack_forget()  # إزالة من التخطيط الحالي

                    if i < half_count:
                        # الصف الأول
                        card.pack(in_=top_counter_frame, side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
                    else:
                        # الصف الثاني
                        card.pack(in_=bottom_counter_frame, side=tk.RIGHT, padx=5, fill=tk.X, expand=True)

    def find_parent_frame(self, widget):
        """العثور على إطار الأب لعنصر واجهة"""
        if widget is None:
            return None

        parent = widget.master
        while parent is not None:
            if isinstance(parent, tk.LabelFrame) and parent.cget("text") == "إحصائيات اليوم":
                return parent
            parent = parent.master

        return None

    def organize_buttons_for_small_screen(self):
        """تنظيم الأزرار للشاشات الصغيرة"""
        # تنفيذ فقط عند الضرورة
        if hasattr(self, 'current_layout') and self.current_layout == "compact":
            return

        # هنا يمكن تنفيذ تغييرات على تنظيم الأزرار
        # مثل إنشاء قائمة منسدلة لبعض الأزرار الأقل استخداماً
        # أو تصغير حجم الأزرار أو تقليل النص المعروض

        pass  # يمكن تنفيذ المزيد حسب الاحتياج

    def show_all_buttons(self):
        """عرض جميع الأزرار للشاشات الكبيرة"""
        # تنفيذ فقط عند الضرورة
        if hasattr(self, 'current_layout') and self.current_layout == "expanded":
            return

        # إعادة الأزرار إلى حالتها الطبيعية
        # مثل إظهار جميع الأزرار وإعادة النصوص الكاملة

        pass  # يمكن تنفيذ المزيد حسب الاحتياج

    def setup_users_tab(self):
        user_management_frame = tk.Frame(self.users_tab, bg=self.colors["light"])
        user_management_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(
            user_management_frame,
            text="إدارة مستخدمي النظام (محمي بكلمة مرور) - خاص بالمشرف",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        open_button = tk.Button(
            user_management_frame,
            text="فتح نافذة إدارة المستخدمين",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=20, pady=10, bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=self.protected_open_user_management
        )
        open_button.pack(pady=50)

        # إضافة إطار للنسخ الاحتياطي
        backup_frame = tk.Frame(user_management_frame, bg=self.colors["light"], pady=20)
        backup_frame.pack(pady=20)

        tk.Label(
            backup_frame,
            text="إدارة النسخ الاحتياطية لقاعدة البيانات",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["dark"]
        ).pack(pady=(0, 10))

        # إضافة أزرار النسخ الاحتياطي والاسترداد
        backup_btn = tk.Button(
            backup_frame,
            text="إنشاء نسخة احتياطية",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=self.backup_database
        )
        backup_btn.pack(side=tk.LEFT, padx=5)

        restore_btn = tk.Button(
            backup_frame,
            text="استرداد نسخة احتياطية",
            font=self.fonts["text_bold"],
            bg=self.colors["warning"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=self.restore_database
        )
        restore_btn.pack(side=tk.LEFT, padx=5)

        tk.Label(
            user_management_frame,
            text="لن يتم فتح نافذة إدارة المستخدمين إلا بعد إدخال كلمة مرور المشرف.",
            font=self.fonts["text"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=10, pady=10, wraplength=700
        ).pack(fill=tk.X)

    def protected_open_user_management(self):
        if not self.current_user["permissions"]["is_admin"]:
            messagebox.showerror("خطأ", "لا تملك صلاحية!")
            return
        admin_pass = simpledialog.askstring("إدخال كلمة المرور", "أدخل كلمة المرور الخاصة بالمشرف:", show='*')
        if not admin_pass:
            return
        cur = self.conn.cursor()
        cur.execute("SELECT password FROM users WHERE username='admin'")
        row = cur.fetchone()
        if not row:
            messagebox.showerror("خطأ", "لا يوجد حساب مشرف رئيسي!")
            return
        admin_real_hash = row[0]
        hashed_input = hashlib.sha256(admin_pass.encode()).hexdigest()
        if hashed_input == admin_real_hash:
            UserManagement(self.root, self.conn, self.current_user, self.colors, self.fonts)
        else:
            messagebox.showerror("خطأ", "كلمة المرور غير صحيحة!")

    def create_tables(self):
        try:
            with self.conn:
                # تعديل جدول الطلاب لإضافة حقول الاستبعاد
                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS trainees (
                        national_id TEXT PRIMARY KEY,
                        name TEXT,
                        rank TEXT,
                        course TEXT,
                        phone TEXT,
                        is_excluded INTEGER DEFAULT 0,
                        exclusion_reason TEXT DEFAULT '',
                        excluded_date TEXT DEFAULT ''
                    )
                """)

                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS attendance (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        national_id TEXT,
                        name TEXT,
                        rank TEXT,
                        course TEXT,
                        time TEXT,
                        date TEXT,
                        status TEXT,
                        original_status TEXT,
                        registered_by TEXT,
                        excuse_reason TEXT DEFAULT '',
                        updated_by TEXT,
                        updated_at TEXT,
                        modification_reason TEXT DEFAULT ''
                    )
                """)

                # إضافة جدول الفصول
                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS course_sections (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        course_name TEXT NOT NULL,
                        section_name TEXT NOT NULL,
                        created_date TEXT,
                        UNIQUE(course_name, section_name)
                    )
                """)

                # إضافة جدول تسجيل الطلاب في الفصول
                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS student_sections (
                        national_id TEXT NOT NULL,
                        course_name TEXT NOT NULL,
                        section_name TEXT NOT NULL,
                        assigned_date TEXT,
                        PRIMARY KEY (national_id, course_name),
                        FOREIGN KEY (national_id) REFERENCES trainees(national_id)
                    )
                """)

                # إضافة جدول معلومات الدورات (إذا لم يكن موجوداً)
                self.conn.execute("""
                               CREATE TABLE IF NOT EXISTS course_info (
                                   course_name TEXT PRIMARY KEY,
                                   start_day TEXT,
                                   start_month TEXT,
                                   start_year TEXT,
                                   end_day TEXT,
                                   end_month TEXT,
                                   end_year TEXT,
                                   created_date TEXT
                               )
                           """)

                # إضافة جدول المخالفات
                self.conn.execute("""
                                               CREATE TABLE IF NOT EXISTS student_violations (
                                                   id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                   national_id TEXT,
                                                   violation_date TEXT,
                                                   violation_type TEXT,
                                                   description TEXT,
                                                   action_taken TEXT,
                                                   action_date TEXT,
                                                   recorded_by TEXT,
                                                   notes TEXT,
                                                   attachment_path TEXT,
                                                   FOREIGN KEY (national_id) REFERENCES trainees(national_id)
                                               )
                                           """)

                # إضافة أعمدة الاستبعاد للطلاب الحاليين إذا لم تكن موجودة
                cursor = self.conn.cursor()
                cursor.execute("PRAGMA table_info(trainees)")
                columns = [column[1] for column in cursor.fetchall()]

                if "is_excluded" not in columns:
                    self.conn.execute("ALTER TABLE trainees ADD COLUMN is_excluded INTEGER DEFAULT 0")
                if "exclusion_reason" not in columns:
                    self.conn.execute("ALTER TABLE trainees ADD COLUMN exclusion_reason TEXT DEFAULT ''")
                if "excluded_date" not in columns:
                    self.conn.execute("ALTER TABLE trainees ADD COLUMN excluded_date TEXT DEFAULT ''")

                # فحص وإضافة الأعمدة المفقودة في جدول attendance
                # فحص وإضافة الأعمدة المفقودة في جدول attendance
                cursor.execute("PRAGMA table_info(attendance)")
                columns = [column[1] for column in cursor.fetchall()]

                # إضافة الأعمدة المفقودة إذا لم تكن موجودة
                if "original_status" not in columns:
                    self.conn.execute("ALTER TABLE attendance ADD COLUMN original_status TEXT")
                if "updated_by" not in columns:
                    self.conn.execute("ALTER TABLE attendance ADD COLUMN updated_by TEXT")
                if "updated_at" not in columns:
                    self.conn.execute("ALTER TABLE attendance ADD COLUMN updated_at TEXT")

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء إنشاء/تعديل الجداول: {str(e)}")

    def create_header(self):
        header_frame = tk.Frame(self.root, bg=self.colors["primary"], height=70)
        header_frame.pack(fill=tk.X)

        # منع الإطار من الانكماش
        header_frame.pack_propagate(False)

        # استخدام تخطيط أكثر مرونة للعناوين
        logo_container = tk.Frame(header_frame, bg=self.colors["primary"])
        logo_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        logo_label = tk.Label(
            logo_container,
            text="قسم شؤون المدربين - نظام الحضور والغياب",
            font=self.fonts["large_title"],
            bg=self.colors["primary"],
            fg="white"
        )
        logo_label.pack(side=tk.RIGHT)

        # إطار منفصل لمعلومات المستخدم
        user_frame = tk.Frame(logo_container, bg=self.colors["primary"])
        user_frame.pack(side=tk.LEFT)

        user_type = "مشرف" if self.current_user["permissions"]["is_admin"] else "مستخدم"
        user_label = tk.Label(
            user_frame,
            text=f"المستخدم: {self.current_user['full_name']} ({user_type})",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white"
        )
        user_label.pack(side=tk.LEFT)

        logout_btn = tk.Button(
            user_frame,
            text="تسجيل الخروج",
            font=self.fonts["text"],
            bg="#ff5252",
            fg="white",
            padx=10,
            pady=2,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.logout
        )
        logout_btn.pack(side=tk.LEFT, padx=15)

        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="ملف", menu=file_menu)
        file_menu.add_command(label="التحقق من التحديثات", command=version_manager.check_for_updates)
        file_menu.add_separator()
        file_menu.add_command(label="خروج", command=self.root.destroy)

        # تعديل حجم العنوان للشاشات الصغيرة
        def adjust_header_for_screen_size(event=None):
            window_width = self.root.winfo_width()
            if window_width < 800:
                logo_label.config(text="نظام الحضور والغياب")
            else:
                logo_label.config(text="قسم شؤون المدربين - نظام الحضور والغياب")

        # ربط وظيفة تغيير الحجم بحدث تغيير حجم النافذة
        self.root.bind('<Configure>', adjust_header_for_screen_size)

    def logout(self):
        if messagebox.askyesnocancel("تسجيل الخروج", "هل تريد تسجيل الخروج؟"):
            self.root.destroy()
            root = tk.Tk()
            LoginSystem(root)
            root.mainloop()

    def filter_attendance(self, event=None):
        selected_status = self.status_filter_var.get()
        if selected_status == "الكل":
            self.export_button.config(text="تصدير الكل")
        else:
            self.export_button.config(text=f"تصدير {selected_status}")
        self.update_attendance_display()

    def get_all_absences_count(self, national_id):
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) 
            FROM attendance
            WHERE national_id=? AND status='غائب'
        """, (national_id,))
        return cursor.fetchone()[0]

    def get_all_late_count(self, national_id):
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) 
            FROM attendance
            WHERE national_id=? AND status='متأخر'
        """, (national_id,))
        return cursor.fetchone()[0]

    def get_all_excused_count(self, national_id):
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) 
            FROM attendance
            WHERE national_id=? AND status='غائب بعذر'
        """, (national_id,))
        return cursor.fetchone()[0]

    def setup_attendance_tab(self):
        self.setup_stats_panel()
        self.setup_individual_attendance()
        self.setup_bulk_attendance()

    def setup_stats_panel(self):
        stats_frame = tk.LabelFrame(
            self.attendance_tab,
            text="إحصائيات اليوم",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=10, pady=10
        )
        stats_frame.pack(fill=tk.X, padx=10, pady=5)

        # إنشاء إطارين للصفين
        top_counter_frame = tk.Frame(stats_frame, bg=self.colors["light"])
        top_counter_frame.pack(fill=tk.X, padx=5, pady=5)

        bottom_counter_frame = tk.Frame(stats_frame, bg=self.colors["light"])
        bottom_counter_frame.pack(fill=tk.X, padx=5, pady=5)

        # مسح قائمة البطاقات
        self.stats_cards = []

        # الصف الأول من الإحصائيات
        card1 = self.create_stat_card(top_counter_frame, "إجمالي الطلاب", self.total_students_var,
                                      self.colors["primary"])
        card1.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card1)

        card2 = self.create_stat_card(top_counter_frame, "الحاضرون", self.present_students_var, self.colors["success"])
        card2.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card2)

        card3 = self.create_stat_card(top_counter_frame, "المتأخرون", self.late_students_var, self.colors["late"])
        card3.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card3)

        card4 = self.create_stat_card(top_counter_frame, "الغائبون", self.absent_students_var, self.colors["danger"])
        card4.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card4)

        card5 = self.create_stat_card(top_counter_frame, "غائب بعذر", self.excused_students_var, self.colors["excused"])
        card5.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card5)

        card6 = self.create_stat_card(top_counter_frame, "لم يباشر", self.not_started_students_var,
                                      self.colors["not_started"])
        card6.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card6)

        # الصف الثاني من الإحصائيات
        card7 = self.create_stat_card(bottom_counter_frame, "تطبيق ميداني", self.field_application_var,
                                      self.colors["field_application"])
        card7.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card7)

        card8 = self.create_stat_card(bottom_counter_frame, "يوم طالب", self.student_day_var,
                                      self.colors["student_day"])
        card8.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card8)

        card9 = self.create_stat_card(bottom_counter_frame, "مسائية / عن بعد", self.evening_remote_var,
                                      self.colors["evening_remote"])
        card9.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card9)

        card10 = self.create_stat_card(bottom_counter_frame, "حالة وفاة", self.death_case_var,
                                       self.colors["death_case"])
        card10.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card10)

        card11 = self.create_stat_card(bottom_counter_frame, "منوم", self.hospital_var, self.colors["hospital"])
        card11.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card11)

        card12 = self.create_stat_card(bottom_counter_frame, "نسبة الحضور", self.attendance_rate_var,
                                       self.colors["warning"])
        card12.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.stats_cards.append(card12)

    def create_stat_card(self, parent, title, variable, color):
        card = tk.Frame(parent, bg=self.colors["light"], bd=1, relief=tk.RIDGE)

        # استخدام نسب مرنة لحجم العنصر
        title_label = tk.Label(card, text=title, font=self.fonts["text_bold"], bg=color, fg="white", padx=5, pady=5)
        title_label.pack(fill=tk.X)

        value_label = tk.Label(card, textvariable=variable, font=self.fonts["title"], bg=self.colors["light"], pady=10)
        value_label.pack(fill=tk.X)

        return card

    def setup_individual_attendance(self):
        """إعادة تصميم إطار تسجيل الحضور بتنظيم أفضل وأكثر راحة للعين"""
        attendance_frame = tk.LabelFrame(
            self.attendance_tab,
            text="تسجيل الحضور",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=15,
            pady=15
        )
        attendance_frame.pack(fill=tk.BOTH, padx=10, pady=5)

        # إنشاء إطار للبحث مع تصميم أفضل
        search_section = tk.Frame(attendance_frame, bg=self.colors["light"])
        search_section.pack(fill=tk.X, pady=(0, 10))

        # تقسيم منطقة البحث إلى قسمين - يمين ويسار
        search_right = tk.Frame(search_section, bg=self.colors["light"])
        search_right.pack(side=tk.RIGHT, fill=tk.Y)

        search_left = tk.Frame(search_section, bg=self.colors["light"])
        search_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # منطقة التاريخ على اليمين
        date_frame = tk.Frame(search_right, bg=self.colors["light"])
        date_frame.pack(side=tk.RIGHT, padx=10, fill=tk.Y)

        date_label = tk.Label(
            date_frame,
            text="التاريخ:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        )
        date_label.pack(side=tk.RIGHT, padx=5)

        self.date_entry = DateEntry(
            date_frame,
            width=12,
            background=self.colors["primary"],
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd',
            font=self.fonts["text"],
            firstweekday="sunday",
            disableddays=(5, 6)
        )
        self.date_entry.pack(side=tk.RIGHT, padx=5)
        self.date_entry.set_date(self.today)
        self.date_entry.bind("<<DateEntrySelected>>", lambda e: self.update_statistics())

        # منطقة البحث على اليسار مع تصميم محسن
        search_box_frame = tk.Frame(search_left, bg=self.colors["light"])
        search_box_frame.pack(fill=tk.X, padx=10)

        search_icon_label = tk.Label(
            search_box_frame,
            text="🔍",
            font=(self.fonts["text"][0], 14),
            bg=self.colors["light"],
            fg=self.colors["primary"]
        )
        search_icon_label.pack(side=tk.RIGHT, padx=(0, 5))

        search_label = tk.Label(
            search_box_frame,
            text="بحث بالاسم أو الهوية:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        )
        search_label.pack(side=tk.RIGHT, padx=5)

        # مربع بحث بحجم أصغر ومظهر أفضل
        self.name_search_entry = tk.Entry(
            search_box_frame,
            font=self.fonts["text"],
            width=20,  # تقليل العرض
            bd=2,
            relief=tk.GROOVE
        )
        self.name_search_entry.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=False)
        self.name_search_entry.bind("<KeyRelease>", self.dynamic_name_search)

        # تحسين مظهر قائمة النتائج
        results_frame = tk.Frame(attendance_frame, bg=self.colors["light"], pady=5)
        results_frame.pack(fill=tk.X, padx=10)

        results_label = tk.Label(
            results_frame,
            text="نتائج البحث:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["primary"]
        )
        results_label.pack(side=tk.RIGHT, anchor=tk.N, padx=(0, 5))

        # إطار لقائمة النتائج مع شريط تمرير
        listbox_frame = tk.Frame(results_frame, bg=self.colors["light"])
        listbox_frame.pack(fill=tk.X, expand=True, side=tk.LEFT)

        # إضافة شريط تمرير
        listbox_scrollbar = tk.Scrollbar(listbox_frame)
        listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.name_listbox = tk.Listbox(
            listbox_frame,
            font=self.fonts["text"],
            height=4,
            selectbackground=self.colors["primary"],
            bd=2,
            relief=tk.GROOVE,
            yscrollcommand=listbox_scrollbar.set
        )
        self.name_listbox.pack(fill=tk.X, expand=True)
        self.name_listbox.bind("<<ListboxSelect>>", self.on_name_select)

        # ربط شريط التمرير بالقائمة
        listbox_scrollbar.config(command=self.name_listbox.yview)

        # حقل خفي لتخزين الهوية
        self.id_entry = tk.Entry(self.root)

        # إضافة فاصل بصري بين البحث وأزرار التسجيل
        separator = tk.Frame(attendance_frame, height=2, bg="#E5E7EB")  # استخدام لون مباشرة
        separator.pack(fill=tk.X, padx=20, pady=10)

        # تحسين تصميم أزرار تسجيل الحضور - استخدام Grid للتوزيع المتساوي
        buttons_frame = tk.Frame(attendance_frame, bg=self.colors["light"])
        buttons_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # توزيع الأعمدة بشكل متساوٍ
        for i in range(5):
            buttons_frame.columnconfigure(i, weight=1)

        # إنشاء إطارات الصفوف
        row1_frame = tk.Frame(buttons_frame, bg=self.colors["light"])
        row1_frame.pack(fill=tk.X, pady=(0, 5))

        row2_frame = tk.Frame(buttons_frame, bg=self.colors["light"])
        row2_frame.pack(fill=tk.X, pady=(5, 0))

        # الصف الأول من الأزرار
        buttons_row1 = [
            ("حاضر", self.colors["success"], lambda: self.insert_attendance_record("حاضر")),
            ("متأخر", self.colors["late"], lambda: self.insert_attendance_record("متأخر")),
            ("غائب", self.colors["danger"], lambda: self.insert_attendance_record("غائب")),
            ("غياب بعذر", self.colors["excused"], self.register_excused_absence),
            ("لم يباشر", self.colors["not_started"], lambda: self.insert_attendance_record("لم يباشر"))
        ]

        # الصف الثاني من الأزرار
        buttons_row2 = [
            ("تطبيق ميداني", self.colors["field_application"], lambda: self.insert_attendance_record("تطبيق ميداني")),
            ("يوم طالب", self.colors["student_day"], lambda: self.insert_attendance_record("يوم طالب")),
            (
            "مسائية / عن بعد", self.colors["evening_remote"], lambda: self.insert_attendance_record("مسائية / عن بعد")),
            ("حالة وفاة", self.colors["death_case"], lambda: self.register_special_case("حالة وفاة")),
            ("منوم", self.colors["hospital"], lambda: self.register_special_case("منوم"))
        ]

        # إنشاء أزرار الصف الأول
        for i, (text, color, command) in enumerate(buttons_row1):
            btn = tk.Button(
                row1_frame,
                text=text,
                font=self.fonts["text_bold"],
                bg=color,
                fg="white",
                padx=5,
                pady=8,
                bd=0,
                relief=tk.FLAT,
                cursor="hand2",
                command=command
            )
            btn.pack(side=tk.RIGHT, padx=3, fill=tk.X, expand=True)

        # إنشاء أزرار الصف الثاني
        for i, (text, color, command) in enumerate(buttons_row2):
            btn = tk.Button(
                row2_frame,
                text=text,
                font=self.fonts["text_bold"],
                bg=color,
                fg="white",
                padx=5,
                pady=8,
                bd=0,
                relief=tk.FLAT,
                cursor="hand2",
                command=command
            )
            btn.pack(side=tk.RIGHT, padx=3, fill=tk.X, expand=True)

        # إضافة مؤشر آخر تسجيل بتصميم محسن
        status_frame = tk.Frame(attendance_frame, bg=self.colors["light"], pady=5)
        status_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

        self.last_registered_label = tk.Label(
            status_frame,
            text="",
            font=self.fonts["text_bold"],
            fg=self.colors["primary"],
            bg=self.colors["light"],
            anchor=tk.W  # محاذاة النص إلى اليمين
        )
        self.last_registered_label.pack(fill=tk.X)

    def get_arabic_day_name(self, date_str):
        """تحويل التاريخ إلى اسم اليوم بالعربية"""
        try:
            # تحويل النص إلى تاريخ
            date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d')

            # الحصول على رقم اليوم في الأسبوع (0=الاثنين، 6=الأحد)
            day_num = date_obj.weekday()

            # قائمة أيام الأسبوع بالعربية (مرتبة حسب نظام Python للأيام)
            arabic_days = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]

            return arabic_days[day_num]
        except:
            # في حال حدوث أي خطأ، نعيد نص التاريخ كما هو
            return date_str


# التحقق من الترخيص قبل بدء البرنامج
def verify_license():
    if not hardware_verification.check_license():
        # إظهار نافذة التفعيل وطلب مفتاح التفعيل
        import tkinter as tk
        from tkinter import messagebox, ttk

        try:
            import pyperclip  # للنسخ واللصق
        except ImportError:
            # تثبيت المكتبة إذا لم تكن موجودة
            import subprocess
            subprocess.check_call(["pip", "install", "pyperclip"])
            import pyperclip

        activation_window = tk.Tk()
        activation_window.title("تفعيل البرنامج")
        activation_window.geometry("500x350")
        activation_window.configure(bg="#f0f0f0")
        # السماح بتكبير النافذة
        activation_window.resizable(True, True)
        activation_window.minsize(400, 300)  # الحد الأدنى للحجم

        # توسيط النافذة
        screen_width = activation_window.winfo_screenwidth()
        screen_height = activation_window.winfo_screenheight()
        x = (screen_width - 500) // 2
        y = (screen_height - 350) // 2
        activation_window.geometry(f"500x350+{x}+{y}")

        # الحصول على معلومات التفعيل
        activation_info = hardware_verification.get_activation_info()
        machine_id = activation_info['machine_id']

        # استخدام ttk.Frame لتحسين المظهر وإضافة padding
        main_frame = ttk.Frame(activation_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # عنوان التفعيل
        ttk.Label(
            main_frame,
            text="تفعيل البرنامج",
            font=("Arial", 18, "bold")
        ).pack(pady=(0, 15))

        # لوحة التفعيل
        activation_frame = ttk.LabelFrame(main_frame, text="معلومات التفعيل", padding="10")
        activation_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # معرف الجهاز
        ttk.Label(
            activation_frame,
            text="معرف الجهاز:",
            font=("Arial", 11, "bold")
        ).pack(anchor=tk.W, pady=(5, 0))

        # إطار لمعرف الجهاز مع زر النسخ
        id_frame = ttk.Frame(activation_frame)
        id_frame.pack(fill=tk.X, pady=5)

        machine_id_var = tk.StringVar(value=machine_id)
        machine_id_entry = ttk.Entry(id_frame, textvariable=machine_id_var, font=("Consolas", 10), width=50)
        machine_id_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        machine_id_entry.config(state="readonly")

        def copy_machine_id():
            try:
                pyperclip.copy(machine_id)
                messagebox.showinfo("تم النسخ", "تم نسخ معرف الجهاز إلى الحافظة")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء النسخ: {str(e)}")

        copy_id_btn = ttk.Button(
            id_frame,
            text="نسخ المعرف",
            command=copy_machine_id
        )
        copy_id_btn.pack(side=tk.RIGHT)

        # رسالة إرشادية
        ttk.Label(
            activation_frame,
            text="أرسل معرف الجهاز للمسؤول للحصول على مفتاح التفعيل",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=(5, 15))

        # مفتاح التفعيل
        ttk.Label(
            activation_frame,
            text="مفتاح التفعيل:",
            font=("Arial", 11, "bold")
        ).pack(anchor=tk.W, pady=(5, 0))

        # إطار لمفتاح التفعيل مع زر اللصق
        key_frame = ttk.Frame(activation_frame)
        key_frame.pack(fill=tk.X, pady=5)

        activation_key_var = tk.StringVar()
        activation_key_entry = ttk.Entry(key_frame, textvariable=activation_key_var, font=("Consolas", 10), width=50)
        activation_key_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        def paste_activation_key():
            try:
                clipboard_text = pyperclip.paste()
                activation_key_var.set(clipboard_text)
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء اللصق: {str(e)}")

        paste_key_btn = ttk.Button(
            key_frame,
            text="لصق",
            command=paste_activation_key
        )
        paste_key_btn.pack(side=tk.RIGHT)

        # إطار للأزرار
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))

        def manual_activate():
            """تفعيل عبر مفتاح مخصص"""
            key = activation_key_var.get().strip()
            if not key:
                messagebox.showerror("خطأ", "يرجى إدخال مفتاح التفعيل")
                return

            # طباعة تشخيصية للتحقق
            print(f"المفتاح المدخل: {key[:20]}...")

            # تعطيل الزر أثناء عملية التفعيل
            activate_btn.state(['disabled'])
            activation_window.update()

            result = hardware_verification.activate_software(key)
            print(f"نتيجة التفعيل: {result}")

            if result:
                messagebox.showinfo("تم التفعيل", "تم تفعيل البرنامج بنجاح!")
                activation_window.destroy()
            else:
                messagebox.showerror("خطأ", "مفتاح التفعيل غير صحيح، يرجى المحاولة مرة أخرى.")
                activate_btn.state(['!disabled'])

        exit_btn = ttk.Button(
            button_frame,
            text="إلغاء",
            command=lambda: exit(0)
        )
        exit_btn.pack(side=tk.LEFT, padx=5)

        activate_btn = ttk.Button(
            button_frame,
            text="تفعيل البرنامج",
            command=manual_activate,
            style="Accent.TButton"  # يمكن استخدام هذا النمط إذا كان متاحًا في إصدار tkinter لديك
        )
        activate_btn.pack(side=tk.RIGHT, padx=5)

        # إضافة اختصارات لوحة المفاتيح
        activation_key_entry.bind("<Control-v>", lambda e: paste_activation_key())
        activation_window.bind("<Return>", lambda e: manual_activate())

        # تركيز حقل مفتاح التفعيل
        activation_key_entry.focus_set()

        # محاولة تطبيق سمة أفضل إذا كان ذلك ممكناً
        try:
            # محاولة إنشاء نمط زر بارز
            style = ttk.Style()
            style.configure("Accent.TButton", font=("Arial", 10, "bold"))
            if hasattr(style, "map"):
                style.map("Accent.TButton",
                          foreground=[("pressed", "white"), ("active", "white")],
                          background=[("pressed", "#2980b9"), ("active", "#3498db")]
                          )
        except Exception:
            pass  # تجاهل الأخطاء في حالة عدم دعم النظام

        # تشغيل النافذة
        activation_window.mainloop()

        # التحقق مرة أخرى بعد النافذة
        if not hardware_verification.check_license():
            exit(0)  # إنهاء البرنامج إذا لم يتم التفعيل



# =============================================================================
#                      نقطة التشغيل الرئيسية
# =============================================================================
if __name__ == "__main__":
    # التحقق من الترخيص قبل بدء البرنامج
    verify_license()

    # البدء في تشغيل البرنامج
    root = tk.Tk()
    LoginSystem(root)
    root.mainloop()
