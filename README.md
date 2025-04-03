import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import datetime
import os
import sqlite3
import pandas as pd
from tkcalendar import DateEntry
import hashlib
import re
# أضف هذا الكود بعد باقي الاستيرادات في بداية الملف
import json
import zipfile
import tempfile
import shutil
import datetime
import uuid


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
            self.conn = sqlite3.connect("attendance.db")

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

    def register_attendance(self, event=None):
        self.insert_attendance_record("حاضر")

    def register_excused_absence(self):
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الغياب بعذر")
            return

        reason = simpledialog.askstring("غياب بعذر", "اكتب سبب الغياب إن وُجد:")
        if reason is None:
            return
        self.insert_attendance_record("غائب بعذر", excuse_reason=reason)

    def insert_attendance_record(self, status, excuse_reason=""):
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        national_id = self.id_entry.get().strip()
        if not national_id:
            messagebox.showwarning("تنبيه", "الرجاء اختيار طالب من خلال البحث بالاسم أو الهوية")
            return

        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT national_id, name, rank, course, is_excluded 
            FROM trainees 
            WHERE national_id=?
        """, (national_id,))

        trainee = cursor.fetchone()
        if not trainee:
            messagebox.showwarning("تنبيه", "لا يوجد طالب بهذا الرقم")
            return

        # التحقق من استبعاد الطالب
        if trainee[4] == 1:
            messagebox.showwarning("تنبيه", "هذا الطالب مستبعد ولا يمكن تسجيل حضوره")
            return

        current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?", (trainee[0], current_date))
        existing_record = cursor.fetchone()

        if existing_record:
            existing_status = existing_record[0]

            # استخدام نوافذ خطأ بدلاً من معلومات لجذب انتباه المستخدم
            if existing_status == status:
                # إذا كانت نفس الحالة
                messagebox.showerror("خطأ في التكرار",
                                     f"⚠️ تنبيه: الطالب {trainee[1]} مسجل بالفعل بحالة '{existing_status}' اليوم\n\nلا يمكن تكرار نفس الحالة للطالب في نفس اليوم.")
            else:
                # إذا كانت حالة مختلفة
                messagebox.showerror("تعارض في الحالة",
                                     f"⚠️ تنبـــيه: الطالب {trainee[1]} مسجل بالفعل بحالة '{existing_status}' اليوم\n\nلتغيير الحالة من '{existing_status}' إلى '{status}'، يرجى استخدام خاصية تعديل الحضور من قائمة سجل الحضور.")

            # مسح قيمة الهوية
            self.id_entry.delete(0, tk.END)
            self.name_search_entry.delete(0, tk.END)
            return

        # التحقق من حالة الطالب في اليوم السابق
        current_date_obj = self.date_entry.get_date()
        yesterday_date_obj = current_date_obj - datetime.timedelta(days=1)
        yesterday_date = yesterday_date_obj.strftime("%Y-%m-%d")

        cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?", (trainee[0], yesterday_date))
        yesterday_record = cursor.fetchone()

        # التحقق إذا كان الطالب مسجل "لم يباشر" في اليوم السابق ويحاول المستخدم تسجيله "غائب"
        if yesterday_record and yesterday_record[0] == "لم يباشر" and status == "غائب":
            response = messagebox.askquestion("تنبيه هام ⚠️",
                                              f"الطالب {trainee[1]} مسجل كـ 'لم يباشر' في اليوم السابق.\n\n"
                                              "• تأكد من أن الطالب لم يباشر الدورة فعلاً.\n"
                                              "• إذا حضر الطالب اليوم، يجب تسجيله كـ 'حاضر' أو 'متأخر'.\n"
                                              "• استمرار تسجيله كـ 'غائب' يعتبر مخالف لتعلميات التدريب المستديمة.\n\n"
                                              "هل تريد تغيير الحالة إلى 'لم يباشر' بدلاً من 'غائب'؟",
                                              icon="warning")
            if response == "yes":
                status = "لم يباشر"
            elif response == "no":
                # إضافة تأكيد إضافي عند الإصرار على الغياب
                confirm = messagebox.askquestion("تأكيد نهائي",
                                                 f"هل أنت متأكد من تسجيل الطالب {trainee[1]} كـ 'غائب' رغم أنه 'لم يباشر' بالأمس؟",
                                                 icon="warning")
                if confirm != "yes":
                    return

        t_id, t_name, t_rank, t_course, _ = trainee
        current_time = datetime.datetime.now().strftime("%H:%M:%S")

        try:
            with self.conn:
                self.conn.execute("""
                    INSERT INTO attendance (
                        national_id, name, rank, course,
                        time, date, status, original_status,
                        registered_by, excuse_reason,
                        updated_by, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    t_id, t_name, t_rank, t_course,
                    current_time, current_date,
                    status, status,
                    self.current_user["full_name"], excuse_reason,
                    "", ""
                ))

            # تحديث رسالة التأكيد في عنصر الواجهة بدلاً من نافذة منبثقة
            if status == "حاضر":
                icon_status = "✅"
            elif status == "غائب":
                icon_status = "❌"
            elif status == "متأخر":
                icon_status = "⏰"
            elif status == "غائب بعذر":
                icon_status = "📝"
            elif status == "لم يباشر":
                icon_status = "⏳"
            else:
                icon_status = "📌"

            # نعرض الرسالة فقط في حقل آخر طالب سُجّل بدلاً من نافذة منبثقة
            self.last_registered_label.config(text=f"آخر طالب سُجِّل: {t_name} ({status}) {icon_status}")

            # مسح حقول الإدخال
            self.id_entry.delete(0, tk.END)
            self.name_search_entry.delete(0, tk.END)
            self.name_listbox.delete(0, tk.END)

            self.update_statistics()
            self.update_attendance_display()
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def setup_bulk_attendance(self):
        """إضافة زر لفتح نافذة إدخال أرقام الهويات بالباركود"""
        # إنشاء إطار بسيط للأزرار الجماعية
        buttons_frame = tk.Frame(self.attendance_tab, bg=self.colors["light"], padx=10, pady=10)
        buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        # عنوان صغير
        tk.Label(
            buttons_frame,
            text="تسجيل الحضور الجماعي:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(side=tk.RIGHT, padx=10, pady=5)

        # زر تسجيل حضور الكل (غير المسجلين)
        bulk_all_present_button = tk.Button(
            buttons_frame,
            text="تسجيل الحضور للكل (غير المسجلين)",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            command=self.register_all_unmarked_as_present,
            cursor="hand2"
        )
        bulk_all_present_button.pack(side=tk.LEFT, padx=5)

        # زر تسجيل حضور دورة كاملة
        course_attendance_button = tk.Button(
            buttons_frame,
            text="تسجيل حضور دورة كاملة",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            command=self.register_whole_course,
            cursor="hand2"
        )
        course_attendance_button.pack(side=tk.LEFT, padx=5)

        # زر فتح نافذة إدخال أرقام الهويات بالباركود
        barcode_window_button = tk.Button(
            buttons_frame,
            text="فتح نافذة إدخال أرقام الهويات بالباركود",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            command=self.open_barcode_window,
            cursor="hand2"
        )
        barcode_window_button.pack(side=tk.LEFT, padx=5)

    def open_barcode_window(self):
        """فتح نافذة جديدة لإدخال أرقام الهويات بالباركود"""
        # إنشاء نافذة جديدة
        barcode_window = tk.Toplevel(self.root)
        barcode_window.title("إدخال أرقام الهويات بالباركود")
        barcode_window.geometry("800x750")  # نافذة أطول لإظهار الأزرار
        barcode_window.configure(bg=self.colors["light"])

        # توسيط النافذة
        x = (barcode_window.winfo_screenwidth() - 800) // 2
        y = (barcode_window.winfo_screenheight() - 750) // 2
        barcode_window.geometry(f"800x750+{x}+{y}")

        # إطار العنوان الرئيسي
        title_frame = tk.Frame(barcode_window, bg=self.colors["primary"], padx=20, pady=20)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame,
            text="إدخال أرقام الهويات بالباركود وتسجيل الحضور الجماعي",
            font=("Tajawal", 24, "bold"),  # خط كبير وعريض
            bg=self.colors["primary"],
            fg="white"
        )
        title_label.pack()

        # إطار التاريخ
        date_frame = tk.Frame(barcode_window, bg=self.colors["light"], padx=20, pady=10)
        date_frame.pack(fill=tk.X)

        tk.Label(
            date_frame,
            text="التاريخ:",
            font=("Tajawal", 16, "bold"),  # خط كبير وعريض
            bg=self.colors["light"]
        ).pack(side=tk.RIGHT, padx=10)

        barcode_date_entry = DateEntry(
            date_frame,
            width=15,
            background=self.colors["primary"],
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd',
            font=("Tajawal", 16),  # خط كبير
            firstweekday="sunday"
        )
        barcode_date_entry.pack(side=tk.RIGHT, padx=10)
        barcode_date_entry.set_date(self.today)

        # إطار إدخال أرقام الهويات
        input_frame = tk.LabelFrame(
            barcode_window,
            text="إدخال أرقام الهويات",
            font=("Tajawal", 18, "bold"),  # خط كبير وعريض
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=20, pady=20
        )
        input_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # توجيهات الاستخدام
        instructions = tk.Label(
            input_frame,
            text="أدخل أرقام الهويات (كل رقم هوية في سطر منفصل):",
            font=("Tajawal", 16, "bold"),  # خط كبير وعريض
            bg=self.colors["light"],
            anchor=tk.W
        )
        instructions.pack(fill=tk.X, pady=(0, 10))

        # إطار النص مع شريط التمرير
        text_frame = tk.Frame(input_frame, bg=self.colors["light"])
        text_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        barcode_text = tk.Text(
            text_frame,
            font=("Tajawal", 14),  # خط أصغر
            height=18,  # زيادة ارتفاع مربع النص
            width=40,
            wrap=tk.WORD,
            yscrollcommand=scrollbar.set,
            bd=2,
            relief=tk.GROOVE
        )
        barcode_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=barcode_text.yview)

        # تركيز مؤشر الكتابة في حقل النص
        barcode_text.focus_set()

        # إطار الأزرار
        buttons_frame = tk.Frame(barcode_window, bg=self.colors["light"], padx=20, pady=20)
        buttons_frame.pack(fill=tk.X)

        # دالة لمعالجة أرقام الهويات
        def process_barcodes(status):
            barcode_ids = barcode_text.get("1.0", tk.END).strip()
            if not barcode_ids:
                messagebox.showinfo("تنبيه", "الرجاء إدخال أرقام الهويات أولاً")
                return

            id_lines = [line.strip() for line in barcode_ids.split("\n") if line.strip()]
            if not id_lines:
                messagebox.showinfo("تنبيه", "لم يتم العثور على أرقام هويات صالحة")
                return

            # الحصول على التاريخ المحدد
            selected_date = barcode_date_entry.get_date().strftime("%Y-%m-%d")
            current_time = datetime.datetime.now().strftime("%H:%M:%S")

            cursor = self.conn.cursor()

            # قوائم لتتبع النتائج
            successful_ids = []
            failed_ids = []
            already_registered_ids = []
            excluded_ids = []

            # معالجة كل رقم هوية
            for national_id in id_lines:
                try:
                    # التحقق من وجود الطالب وما إذا كان مستبعدًا
                    cursor.execute("""
                        SELECT national_id, name, rank, course, is_excluded 
                        FROM trainees 
                        WHERE national_id=?
                    """, (national_id,))

                    trainee = cursor.fetchone()
                    if not trainee:
                        failed_ids.append(national_id)
                        continue

                    # التحقق من استبعاد الطالب
                    if trainee[4] == 1:
                        excluded_ids.append(national_id)
                        continue

                    # التحقق مما إذا كان الطالب مسجلاً بالفعل لهذا اليوم
                    cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?",
                                   (trainee[0], selected_date))
                    existing_record = cursor.fetchone()

                    if existing_record:
                        already_registered_ids.append(national_id)
                        continue

                    # إدراج سجل حضور جديد
                    with self.conn:
                        self.conn.execute("""
                            INSERT INTO attendance (
                                national_id, name, rank, course,
                                time, date, status, original_status,
                                registered_by, excuse_reason,
                                updated_by, updated_at
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            trainee[0], trainee[1], trainee[2], trainee[3],
                            current_time, selected_date,
                            status, status,
                            self.current_user["full_name"], "",
                            "", ""
                        ))

                    successful_ids.append(national_id)

                except Exception as e:
                    print(f"خطأ في معالجة الهوية {national_id}: {str(e)}")
                    failed_ids.append(national_id)

            # إعداد رسالة ملخص النتائج
            result_message = f"تمت معالجة {len(id_lines)} رقم هوية:\n\n"

            if successful_ids:
                result_message += f"✅ تم تسجيل {len(successful_ids)} طالب بنجاح بحالة '{status}'.\n"

            if already_registered_ids:
                result_message += f"⚠️ {len(already_registered_ids)} طالب مسجل مسبقاً في هذا اليوم.\n"

            if excluded_ids:
                result_message += f"❌ {len(excluded_ids)} طالب مستبعد لا يمكن تسجيل حضورهم.\n"

            if failed_ids:
                result_message += f"❓ {len(failed_ids)} رقم هوية غير موجود في قاعدة البيانات."

            # عرض النتائج
            messagebox.showinfo("نتائج تسجيل الحضور", result_message)

            # تفريغ مربع النص بعد المعالجة الناجحة إذا تم تسجيل طلاب بنجاح
            if successful_ids:
                barcode_text.delete("1.0", tk.END)

            # تحديث الإحصائيات وعرض الحضور
            self.update_statistics()
            self.update_attendance_display()

        # أزرار تسجيل الحضور - بحجم أصغر وبدون زر الغياب
        present_btn = tk.Button(
            buttons_frame,
            text="تسجيل حضور",
            font=("Tajawal", 14, "bold"),  # خط أصغر
            bg=self.colors["success"],
            fg="white",
            padx=20, pady=8,  # أحجام أصغر
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: process_barcodes("حاضر")
        )
        present_btn.pack(side=tk.LEFT, padx=10)

        late_btn = tk.Button(
            buttons_frame,
            text="تسجيل تأخير",
            font=("Tajawal", 14, "bold"),  # خط أصغر
            bg=self.colors["late"],
            fg="white",
            padx=20, pady=8,  # أحجام أصغر
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: process_barcodes("متأخر")
        )
        late_btn.pack(side=tk.LEFT, padx=10)

        clear_btn = tk.Button(
            buttons_frame,
            text="تفريغ الحقل",
            font=("Tajawal", 14, "bold"),  # خط أصغر
            bg=self.colors["dark"],
            fg="white",
            padx=20, pady=8,  # أحجام أصغر
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: barcode_text.delete("1.0", tk.END)
        )
        clear_btn.pack(side=tk.RIGHT, padx=10)

        close_btn = tk.Button(
            buttons_frame,
            text="إغلاق",
            font=("Tajawal", 14, "bold"),  # خط أصغر
            bg="#9E9E9E",  # لون رمادي
            fg="white",
            padx=20, pady=8,  # أحجام أصغر
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=barcode_window.destroy
        )
        close_btn.pack(side=tk.RIGHT, padx=10)

    def process_barcode_ids(self, status):
        """معالجة أرقام الهويات المدخلة بالباركود وتسجيل حضورهم"""
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        # قراءة النص من مربع الإدخال
        barcode_text = self.barcode_text.get(1.0, tk.END).strip()
        if not barcode_text:
            messagebox.showinfo("تنبيه", "الرجاء إدخال أرقام الهويات أولاً")
            return

        # تقسيم النص إلى أسطر للحصول على أرقام الهويات
        id_lines = [line.strip() for line in barcode_text.split("\n") if line.strip()]
        if not id_lines:
            messagebox.showinfo("تنبيه", "لم يتم العثور على أرقام هويات صالحة")
            return

        # الحصول على التاريخ الحالي ووقت التسجيل
        current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        current_time = datetime.datetime.now().strftime("%H:%M:%S")

        cursor = self.conn.cursor()

        # قوائم لتتبع النتائج
        successful_ids = []
        failed_ids = []
        already_registered_ids = []
        excluded_ids = []

        # معالجة كل رقم هوية
        for national_id in id_lines:
            # تخطي القيم الفارغة
            if not national_id:
                continue

            try:
                # التحقق من وجود الطالب وما إذا كان مستبعدًا
                cursor.execute("""
                    SELECT national_id, name, rank, course, is_excluded 
                    FROM trainees 
                    WHERE national_id=?
                """, (national_id,))

                trainee = cursor.fetchone()
                if not trainee:
                    failed_ids.append(national_id)
                    continue

                # التحقق من استبعاد الطالب
                if trainee[4] == 1:
                    excluded_ids.append(national_id)
                    continue

                # التحقق مما إذا كان الطالب مسجلاً بالفعل لهذا اليوم
                cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?",
                               (trainee[0], current_date))
                existing_record = cursor.fetchone()

                if existing_record:
                    already_registered_ids.append(national_id)
                    continue

                # إدراج سجل حضور جديد
                with self.conn:
                    self.conn.execute("""
                        INSERT INTO attendance (
                            national_id, name, rank, course,
                            time, date, status, original_status,
                            registered_by, excuse_reason,
                            updated_by, updated_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        trainee[0], trainee[1], trainee[2], trainee[3],
                        current_time, current_date,
                        status, status,
                        self.current_user["full_name"], "",
                        "", ""
                    ))

                successful_ids.append(national_id)

            except Exception as e:
                print(f"خطأ في معالجة الهوية {national_id}: {str(e)}")
                failed_ids.append(national_id)

        # إعداد رسالة ملخص النتائج
        result_message = f"تمت معالجة {len(id_lines)} رقم هوية:\n\n"

        if successful_ids:
            result_message += f"✅ تم تسجيل {len(successful_ids)} طالب بنجاح بحالة '{status}'.\n"

        if already_registered_ids:
            result_message += f"⚠️ {len(already_registered_ids)} طالب مسجل مسبقاً في هذا اليوم.\n"

        if excluded_ids:
            result_message += f"❌ {len(excluded_ids)} طالب مستبعد لا يمكن تسجيل حضورهم.\n"

        if failed_ids:
            result_message += f"❓ {len(failed_ids)} رقم هوية غير موجود في قاعدة البيانات."

        # عرض النتائج
        messagebox.showinfo("نتائج تسجيل الحضور", result_message)

        # تفريغ مربع النص بعد المعالجة الناجحة إذا تم تسجيل طلاب بنجاح
        if successful_ids:
            self.barcode_text.delete(1.0, tk.END)

        # تحديث الإحصائيات وعرض الحضور
        self.update_statistics()
        self.update_attendance_display()

    def register_all_unmarked_as_present(self):
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        current_time = datetime.datetime.now().strftime("%H:%M:%S")

        cursor = self.conn.cursor()
        # استثناء الطلاب المستبعدين
        cursor.execute("""
            SELECT national_id, name, rank, course 
            FROM trainees 
            WHERE is_excluded=0
        """)
        all_students = cursor.fetchall()
        cursor.execute("SELECT DISTINCT national_id FROM attendance WHERE date=?", (current_date,))
        already_recorded = set(row[0] for row in cursor.fetchall())

        new_rows = []
        for student in all_students:
            if student[0] not in already_recorded:
                new_rows.append((
                    student[0], student[1], student[2], student[3],
                    current_time, current_date,
                    "حاضر", "حاضر",
                    self.current_user["full_name"], "",
                    "", ""
                ))

        if not new_rows:
            messagebox.showinfo("ملاحظة", "لا يوجد طلاب غير مسجلين اليوم.")
            return

        try:
            with self.conn:
                self.conn.executemany("""
                    INSERT INTO attendance (
                        national_id, name, rank, course,
                        time, date, status, original_status,
                        registered_by, excuse_reason,
                        updated_by, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, new_rows)
            messagebox.showinfo("نجاح", f"تم تسجيل حضور {len(new_rows)} طالب لم يكونوا مسجلين مسبقًا.")
            self.update_statistics()
            self.update_attendance_display()
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def register_bulk_lateness(self):
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        if not messagebox.askyesnocancel("تأكيد", "هل تريد تسجيل تأخير جماعي لجميع الطلاب الذين لم يتم تسجيلهم اليوم؟"):
            return

        current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        current_time = datetime.datetime.now().strftime("%H:%M:%S")

        cursor = self.conn.cursor()
        # استثناء الطلاب المستبعدين
        cursor.execute("""
            SELECT national_id, name, rank, course 
            FROM trainees 
            WHERE is_excluded=0
        """)
        all_students = cursor.fetchall()
        cursor.execute("SELECT DISTINCT national_id FROM attendance WHERE date=?", (current_date,))
        already_recorded = set(row[0] for row in cursor.fetchall())

        new_late_rows = []
        for student in all_students:
            if student[0] not in already_recorded:
                new_late_rows.append((
                    student[0], student[1], student[2], student[3],
                    current_time, current_date,
                    "متأخر", "متأخر",
                    self.current_user["full_name"], "",
                    "", ""
                ))

        if not new_late_rows:
            messagebox.showinfo("ملاحظة", "لا يوجد طلاب غير مسجلين.")
            return

        try:
            with self.conn:
                self.conn.executemany("""
                    INSERT INTO attendance (
                        national_id, name, rank, course,
                        time, date, status, original_status,
                        registered_by, excuse_reason,
                        updated_by, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, new_late_rows)
            messagebox.showinfo("نجاح", f"تم تسجيل تأخير {len(new_late_rows)} طالب بنجاح")
            self.update_statistics()
            self.update_attendance_display()
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def register_special_case(self, status):
        """دالة تسجيل الحالات الخاصة (وفاة، منوم) مع طلب تفاصيل إضافية"""
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        details = simpledialog.askstring(f"تفاصيل {status}", f"أدخل تفاصيل {status}:")
        if details is None:  # إذا ضغط المستخدم على زر الإلغاء
            return

        self.insert_attendance_record(status, excuse_reason=details)

    def bulk_register(self):
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        text_data = self.bulk_text.get("1.0", tk.END).strip()
        if not text_data:
            return

        current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        lines = text_data.split("\n")

        cursor = self.conn.cursor()
        new_rows = []

        for line in lines:
            nid = line.strip()
            if not nid:
                continue

            # التحقق من استبعاد الطالب
            cursor.execute("""
                SELECT national_id, name, rank, course, is_excluded 
                FROM trainees 
                WHERE national_id=?
            """, (nid,))

            trainee = cursor.fetchone()
            if not trainee:
                continue

            # تخطي الطلاب المستبعدين
            if trainee[4] == 1:
                continue

            cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?", (trainee[0], current_date))
            existing_record = cursor.fetchone()
            if existing_record:
                continue

            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            new_rows.append((
                trainee[0], trainee[1], trainee[2], trainee[3],
                current_time, current_date,
                "حاضر", "حاضر",
                self.current_user["full_name"], "",
                "", ""
            ))

        if not new_rows:
            messagebox.showinfo("ملاحظة", "لا يوجد طلاب جدد للتسجيل.")
            return

        try:
            with self.conn:
                self.conn.executemany("""
                    INSERT INTO attendance (
                        national_id, name, rank, course,
                        time, date, status, original_status,
                        registered_by, excuse_reason,
                        updated_by, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, new_rows)
            messagebox.showinfo("نجاح", f"تم تسجيل حضور {len(new_rows)} طالب بنجاح")
            self.bulk_text.delete("1.0", tk.END)
            self.update_statistics()
            self.update_attendance_display()
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def register_whole_course(self):
        """تسجيل حضور دورة كاملة بحالة محددة"""
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        # إنشاء نافذة اختيار الدورة والحالة
        select_window = tk.Toplevel(self.root)
        select_window.title("تسجيل حضور دورة كاملة")
        select_window.geometry("400x300")
        select_window.configure(bg=self.colors["light"])
        select_window.transient(self.root)
        select_window.grab_set()

        x = (select_window.winfo_screenwidth() - 400) // 2
        y = (select_window.winfo_screenheight() - 300) // 2
        select_window.geometry(f"400x300+{x}+{y}")

        # عنوان النافذة
        title_label = tk.Label(
            select_window,
            text="تسجيل حضور دورة كاملة",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        )
        title_label.pack(fill=tk.X)

        form_frame = tk.Frame(select_window, bg=self.colors["light"], padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)

        # اختيار الدورة
        tk.Label(
            form_frame,
            text="اختر الدورة:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            anchor=tk.E
        ).grid(row=0, column=1, padx=5, pady=10, sticky=tk.E)

        # الحصول على قائمة الدورات الحالية
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course FROM trainees WHERE is_excluded=0")
        courses = [row[0] for row in cursor.fetchall() if row[0]]

        course_var = tk.StringVar()
        course_combobox = ttk.Combobox(
            form_frame,
            textvariable=course_var,
            values=courses,
            state="readonly",
            width=25,
            font=self.fonts["text"]
        )
        course_combobox.grid(row=0, column=0, padx=5, pady=10, sticky=tk.W)

        # اختيار الحالة
        tk.Label(
            form_frame,
            text="اختر الحالة:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            anchor=tk.E
        ).grid(row=1, column=1, padx=5, pady=10, sticky=tk.E)

        status_var = tk.StringVar()
        status_combobox = ttk.Combobox(
            form_frame,
            textvariable=status_var,
            values=["حاضر", "متأخر", "غائب", "لم يباشر", "تطبيق ميداني", "يوم طالب", "مسائية / عن بعد"],
            state="readonly",
            width=25,
            font=self.fonts["text"]
        )
        status_combobox.grid(row=1, column=0, padx=5, pady=10, sticky=tk.W)
        status_combobox.current(0)  # اختيار "حاضر" كقيمة افتراضية

        # أزرار التنفيذ والإلغاء
        button_frame = tk.Frame(select_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X)

        def execute_registration():
            selected_course = course_var.get()
            selected_status = status_var.get()

            if not selected_course:
                messagebox.showwarning("تنبيه", "الرجاء اختيار دورة")
                return

            # الحصول على جميع الطلاب في الدورة المحددة
            cursor.execute("""
                SELECT national_id, name, rank, course 
                FROM trainees 
                WHERE course=? AND is_excluded=0
            """, (selected_course,))
            students = cursor.fetchall()

            if not students:
                messagebox.showinfo("تنبيه", f"لا يوجد طلاب في الدورة {selected_course}")
                return

            current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
            current_time = datetime.datetime.now().strftime("%H:%M:%S")

            # التحقق من الطلاب المسجلين مسبقًا
            cursor.execute("""
                SELECT a.national_id, a.status
                FROM attendance a
                JOIN trainees t ON a.national_id = t.national_id
                WHERE a.date=? AND t.course=? AND t.is_excluded=0
            """, (current_date, selected_course))

            already_registered = {row[0]: row[1] for row in cursor.fetchall()}

            # حساب عدد الطلاب المراد تسجيلهم والمسجلين بالفعل
            total_students = len(students)
            already_registered_count = len(already_registered)
            to_register_count = total_students - already_registered_count

            # استفسار المستخدم إذا كان هناك طلاب مسجلين بالفعل
            if already_registered_count > 0:
                message = f"هناك {already_registered_count} طالب مسجلين بالفعل من أصل {total_students}.\n\n"

                if to_register_count > 0:
                    message += f"هل تريد تسجيل الـ {to_register_count} المتبقين بحالة '{selected_status}'؟"
                    if not messagebox.askyesno("تأكيد", message):
                        return
                else:
                    messagebox.showinfo("تنبيه", "جميع طلاب الدورة مسجلين بالفعل.")
                    return

            # إعداد البيانات للإدخال
            new_records = []
            for student in students:
                student_id, student_name, student_rank, student_course = student

                # تخطي الطلاب المسجلين بالفعل
                if student_id in already_registered:
                    continue

                new_records.append((
                    student_id, student_name, student_rank, student_course,
                    current_time, current_date,
                    selected_status, selected_status,
                    self.current_user["full_name"], "",
                    "", ""
                ))

            # تنفيذ الإدخال الجماعي
            try:
                with self.conn:
                    self.conn.executemany("""
                        INSERT INTO attendance (
                            national_id, name, rank, course,
                            time, date, status, original_status,
                            registered_by, excuse_reason,
                            updated_by, updated_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, new_records)

                messagebox.showinfo("نجاح",
                                    f"تم تسجيل {len(new_records)} طالب من دورة '{selected_course}' بحالة '{selected_status}'")
                select_window.destroy()
                self.update_statistics()
                self.update_attendance_display()
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء تسجيل الحضور: {str(e)}")

        register_btn = tk.Button(
            button_frame,
            text="تسجيل الحضور",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=execute_registration
        )
        register_btn.pack(side=tk.LEFT, padx=20)

        cancel_btn = tk.Button(
            button_frame,
            text="إلغاء",
            font=self.fonts["text_bold"],
            bg=self.colors["danger"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=select_window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=20)

    def dynamic_name_search(self, event):
        try:
            text = self.name_search_entry.get().strip()
            self.name_listbox.delete(0, tk.END)
            if not text:
                return
            cursor = self.conn.cursor()
            # البحث بالاسم أو برقم الهوية معًا
            cursor.execute("""
                SELECT name, national_id 
                FROM trainees 
                WHERE (name LIKE ? OR national_id LIKE ?) AND is_excluded=0
            """, ('%' + text + '%', '%' + text + '%',))

            results = cursor.fetchall()
            for row in results:
                self.name_listbox.insert(tk.END, f"{row[0]} ({row[1]})")
        except (tk.TclError, AttributeError):
            # تجاهل الخطأ إذا لم يعد العنصر موجوداً
            pass

    def on_name_select(self, event):
        selection = self.name_listbox.curselection()
        if not selection:
            return
        selected_text = self.name_listbox.get(selection[0])
        # استخراج رقم الهوية من النص المحدد (الاسم (الهوية))
        try:
            national_id = selected_text.split("(")[1].split(")")[0]
            # تخزين رقم الهوية في الحقل الخفي
            self.id_entry.delete(0, tk.END)
            self.id_entry.insert(0, national_id)
        except:
            pass  # في حال حدوث خطأ في تنسيق النص

    def setup_attendance_log_tab(self):
        """تعديل دالة إعداد تبويب سجل الحضور لجعل الأزرار مرنة"""
        table_frame = tk.LabelFrame(self.attendance_log_tab, text="سجل الحضور", font=self.fonts["subtitle"],
                                    bg=self.colors["light"], fg=self.colors["dark"], padx=10, pady=10)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # إنشاء إطارين منفصلين للتحكم (العلوي والسفلي)
        # الإطار العلوي للبحث والتاريخ والتصفية
        top_controls = tk.Frame(table_frame, bg=self.colors["light"])
        top_controls.pack(fill=tk.X, pady=(5, 2))

        # الإطار السفلي للأزرار
        button_controls = tk.Frame(table_frame, bg=self.colors["light"])
        button_controls.pack(fill=tk.X, pady=(2, 5))

        # توزيع عناصر التحكم بين الإطارين

        # ---------- الإطار العلوي ----------
        # إطار اليمين - التاريخ والتصفية
        right_frame = tk.Frame(top_controls, bg=self.colors["light"])
        right_frame.pack(side=tk.RIGHT, fill=tk.Y)

        tk.Label(right_frame, text="التاريخ:", font=self.fonts["text_bold"], bg=self.colors["light"]).pack(
            side=tk.RIGHT, padx=5)
        self.log_date_entry = DateEntry(
            right_frame,
            width=15,
            background=self.colors["primary"],
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd',
            font=self.fonts["text"],
            firstweekday="sunday",
            disableddays=(5, 6)
        )
        self.log_date_entry.pack(side=tk.RIGHT, padx=5)
        self.log_date_entry.set_date(self.today)
        self.log_date_entry.bind("<<DateEntrySelected>>", lambda e: self.update_attendance_display())

        tk.Label(right_frame, text="تصفية حسب الحالة:", font=self.fonts["text"], bg=self.colors["light"]).pack(
            side=tk.RIGHT, padx=5)

        self.status_filter_var = tk.StringVar()
        self.status_filter = ttk.Combobox(
            right_frame,
            textvariable=self.status_filter_var,
            values=["الكل", "حاضر", "متأخر", "غائب", "غائب بعذر", "لم يباشر",
                    "تطبيق ميداني", "يوم طالب", "مسائية / عن بعد", "حالة وفاة", "منوم"],
            state="readonly",
            width=15,
            font=self.fonts["text"]
        )
        self.status_filter.current(0)
        self.status_filter.pack(side=tk.RIGHT, padx=5)
        self.status_filter.bind("<<ComboboxSelected>>", self.filter_attendance)

        # إطار اليسار - البحث
        left_frame = tk.Frame(top_controls, bg=self.colors["light"])
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(left_frame, text="بحث (الاسم/الهوية):", font=self.fonts["text"], bg=self.colors["light"]).pack(
            side=tk.LEFT, padx=5)
        self.log_search_var = tk.StringVar()
        self.log_search_entry = tk.Entry(left_frame, textvariable=self.log_search_var, font=self.fonts["text"],
                                         width=20)
        self.log_search_entry.pack(side=tk.LEFT, padx=5)
        self.log_search_entry.bind("<KeyRelease>", lambda e: self.update_attendance_display())

        # ---------- الإطار السفلي (للأزرار) ----------
        # استخدام Grid لتوزيع الأزرار بشكل مرن
        button_controls.columnconfigure(0, weight=1)  # للمساحة على اليمين
        button_controls.columnconfigure(1, weight=0)  # للزر الأول
        button_controls.columnconfigure(2, weight=0)  # للزر الثاني
        button_controls.columnconfigure(3, weight=0)  # للزر الثالث
        button_controls.columnconfigure(4, weight=0)  # للزر الرابع
        button_controls.columnconfigure(5, weight=1)  # للمساحة على اليسار

        # إضافة زر إحصائيات الغياب والتأخير
        col_index = 1  # نبدأ من العمود 1

        # إضافة زر أعلى معدلات الغياب والتأخير
        top_absence_button = tk.Button(
            button_controls,
            text="أعلى معدلات الغياب والتأخير",
            font=self.fonts["text_bold"],
            bg="#673AB7",  # لون بنفسجي مميز
            fg="white",
            padx=10,
            pady=3,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.show_top_absence_statistics
        )
        top_absence_button.grid(row=0, column=col_index, padx=5, pady=5, sticky="ew")
        col_index += 1

        # إضافة زر التصدير إذا كان المستخدم لديه صلاحية
        if self.current_user["permissions"]["can_export_data"]:
            self.export_button = tk.Button(
                button_controls,
                text="تصدير الكل",
                font=self.fonts["text_bold"],
                bg=self.colors["primary"],
                fg="white",
                padx=10,
                pady=3,
                bd=0,
                relief=tk.FLAT,
                cursor="hand2",
                command=self.export_based_on_filter
            )
            self.export_button.grid(row=0, column=col_index, padx=5, pady=5, sticky="ew")
            col_index += 1

        # إضافة زر تصدير تكميل الدورات (متاح لجميع المستخدمين)
        completion_export_button = tk.Button(
            button_controls,
            text="تصدير تكميل الدورات",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=10,
            pady=3,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.export_course_completion
        )
        completion_export_button.grid(row=0, column=col_index, padx=5, pady=5, sticky="ew")
        col_index += 1

        # إضافة زر إعادة تعيين السجل إذا كان المستخدم لديه صلاحية
        if self.current_user["permissions"]["can_reset_attendance"]:
            reset_button = tk.Button(
                button_controls,
                text="إعادة تعيين السجل",
                font=self.fonts["text_bold"],
                bg=self.colors["dark"],
                fg="white",
                padx=5,
                pady=3,
                bd=0,
                relief=tk.FLAT,
                command=self.reset_attendance,
                cursor="hand2"
            )
            reset_button.grid(row=0, column=col_index, padx=5, pady=5, sticky="ew")

        # تابع باقي الكود كما هو...
        self.tree_scroll = tk.Scrollbar(table_frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        if self.current_user["permissions"]["can_view_edit_history"]:
            columns = (
                "id",
                "name",
                "rank",
                "course",
                "status",
                "absences",
                "late_count",
                "excused",
                "updated_by",
                "updated_at"
            )
        else:
            columns = (
                "id", "name", "rank", "course", "status",
                "absences", "late_count", "excused"
            )

        self.attendance_tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                            yscrollcommand=self.tree_scroll.set, style="Bold.Treeview")

        self.attendance_tree.column("id", width=100, anchor=tk.CENTER)
        self.attendance_tree.column("name", width=120, anchor=tk.CENTER)
        self.attendance_tree.column("rank", width=80, anchor=tk.CENTER)
        self.attendance_tree.column("course", width=120, anchor=tk.CENTER)
        self.attendance_tree.column("status", width=90, anchor=tk.CENTER)
        self.attendance_tree.column("absences", width=80, anchor=tk.CENTER)
        self.attendance_tree.column("late_count", width=110, anchor=tk.CENTER)
        self.attendance_tree.column("excused", width=90, anchor=tk.CENTER)

        if self.current_user["permissions"]["can_view_edit_history"]:
            self.attendance_tree.column("updated_by", width=120, anchor=tk.CENTER)
            self.attendance_tree.column("updated_at", width=130, anchor=tk.CENTER)

        self.attendance_tree.heading("id", text="رقم الهوية")
        self.attendance_tree.heading("name", text="الاسم")
        self.attendance_tree.heading("rank", text="الرتبة")
        self.attendance_tree.heading("course", text="الدورة")
        self.attendance_tree.heading("status", text="الحالة")
        self.attendance_tree.heading("absences", text="غياب بدون عذر")
        self.attendance_tree.heading("late_count", text="عدد مرات التأخير")
        self.attendance_tree.heading("excused", text="غياب بعذر")

        if self.current_user["permissions"]["can_view_edit_history"]:
            self.attendance_tree.heading("updated_by", text="من عدّل")
            self.attendance_tree.heading("updated_at", text="وقت آخر تعديل")

        self.attendance_tree.pack(fill=tk.BOTH, expand=True)
        self.tree_scroll.config(command=self.attendance_tree.yview)

        self.attendance_tree.tag_configure("present", background="#e8f5e9")
        self.attendance_tree.tag_configure("late", background="#fff8e1")
        self.attendance_tree.tag_configure("excused", background="#e1f5fe")
        self.attendance_tree.tag_configure("not_started", background="#FFE5CC")
        self.attendance_tree.tag_configure("field_application", background="#E0E0E0")  # رمادي فاتح
        self.attendance_tree.tag_configure("student_day", background="#ECECEC")  # رمادي أفتح
        self.attendance_tree.tag_configure("evening_remote", background="#DDDDDD")  # رمادي متوسط
        self.attendance_tree.tag_configure("death_case", background="#E0D6F5")  # لون فاتح للوفاة
        self.attendance_tree.tag_configure("hospital", background="#D4F0ED")  # لون فاتح للمنوم

        if self.current_user["permissions"]["can_edit_attendance"]:
            self.attendance_tree.bind("<Double-1>", self.on_attendance_double_click)

    def show_top_absence_statistics(self):
        """عرض الطلاب الأكثر غياباً وتأخراً وغياباً بعذر"""
        stats_window = tk.Toplevel(self.root)
        stats_window.title("إحصائيات أعلى معدلات الغياب والتأخير")
        stats_window.geometry("900x650")
        stats_window.configure(bg=self.colors["light"])
        stats_window.transient(self.root)
        stats_window.grab_set()

        # توسيط النافذة
        x = (stats_window.winfo_screenwidth() - 900) // 2
        y = (stats_window.winfo_screenheight() - 650) // 2
        stats_window.geometry(f"900x650+{x}+{y}")

        # إطار العنوان
        header_frame = tk.Frame(stats_window, bg=self.colors["primary"], padx=20, pady=15)
        header_frame.pack(fill=tk.X)

        header_label = tk.Label(
            header_frame,
            text="إحصائيات أعلى معدلات الغياب والتأخير",
            font=("Tajawal", 18, "bold"),
            bg=self.colors["primary"],
            fg="white"
        )
        header_label.pack()

        # إطار التحكم
        control_frame = tk.Frame(stats_window, bg=self.colors["light"], padx=20, pady=10)
        control_frame.pack(fill=tk.X)

        tk.Label(
            control_frame,
            text="عدد الطلاب المراد عرضهم:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(side=tk.RIGHT, padx=5)

        limit_var = tk.StringVar(value="10")
        limit_combobox = ttk.Combobox(
            control_frame,
            textvariable=limit_var,
            values=["5", "10", "15", "20", "25", "50"],
            state="readonly",
            width=5,
            font=self.fonts["text"]
        )
        limit_combobox.pack(side=tk.RIGHT, padx=5)

        tk.Label(
            control_frame,
            text="تصفية حسب الدورة:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(side=tk.RIGHT, padx=5)

        # الحصول على قائمة الدورات
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course FROM trainees WHERE course IS NOT NULL AND course != '' ORDER BY course")
        courses = ["جميع الدورات"] + [row[0] for row in cursor.fetchall()]

        course_var = tk.StringVar(value="جميع الدورات")
        course_combobox = ttk.Combobox(
            control_frame,
            textvariable=course_var,
            values=courses,
            state="readonly",
            width=20,
            font=self.fonts["text"]
        )
        course_combobox.pack(side=tk.RIGHT, padx=5)

        refresh_button = tk.Button(
            control_frame,
            text="تحديث البيانات",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: load_statistics()
        )
        refresh_button.pack(side=tk.LEFT, padx=5)

        export_button = tk.Button(
            control_frame,
            text="تصدير البيانات",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: export_statistics()
        )
        export_button.pack(side=tk.LEFT, padx=5)

        # إطار التبويب الرئيسي
        tab_control = ttk.Notebook(stats_window)
        tab_control.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # إنشاء التبويبات الثلاثة
        tab_absence = tk.Frame(tab_control, bg=self.colors["light"])
        tab_lateness = tk.Frame(tab_control, bg=self.colors["light"])
        tab_excused = tk.Frame(tab_control, bg=self.colors["light"])

        tab_control.add(tab_absence, text="أعلى معدلات الغياب")
        tab_control.add(tab_lateness, text="أعلى معدلات التأخير")
        tab_control.add(tab_excused, text="أعلى معدلات الغياب بعذر")

        # إنشاء جداول العرض
        absence_tree = self.create_stats_tree(tab_absence)
        lateness_tree = self.create_stats_tree(tab_lateness)
        excused_tree = self.create_stats_tree(tab_excused)

        # دالة تحميل البيانات
        def load_statistics():
            limit = int(limit_var.get())
            course_filter = course_var.get()

            # مسح البيانات الحالية
            for tree in [absence_tree, lateness_tree, excused_tree]:
                for item in tree.get_children():
                    tree.delete(item)

            # إعداد شرط الدورة
            course_condition = ""
            course_params = []

            if course_filter != "جميع الدورات":
                course_condition = "AND t.course = ?"
                course_params = [course_filter]

            # استعلام الغياب
            query_absence = f"""
            SELECT t.national_id, t.name, t.rank, t.course, COUNT(*) as count
            FROM attendance a
            JOIN trainees t ON a.national_id = t.national_id
            WHERE a.status = 'غائب' AND t.is_excluded = 0 {course_condition}
            GROUP BY t.national_id
            ORDER BY count DESC
            LIMIT ?
            """

            # استعلام التأخير
            query_lateness = f"""
            SELECT t.national_id, t.name, t.rank, t.course, COUNT(*) as count
            FROM attendance a
            JOIN trainees t ON a.national_id = t.national_id
            WHERE a.status = 'متأخر' AND t.is_excluded = 0 {course_condition}
            GROUP BY t.national_id
            ORDER BY count DESC
            LIMIT ?
            """

            # استعلام الغياب بعذر
            query_excused = f"""
            SELECT t.national_id, t.name, t.rank, t.course, COUNT(*) as count
            FROM attendance a
            JOIN trainees t ON a.national_id = t.national_id
            WHERE a.status = 'غائب بعذر' AND t.is_excluded = 0 {course_condition}
            GROUP BY t.national_id
            ORDER BY count DESC
            LIMIT ?
            """

            # تنفيذ الاستعلامات وعرض البيانات
            cursor = self.conn.cursor()

            # الغياب
            cursor.execute(query_absence, course_params + [limit])
            result_absence = cursor.fetchall()
            for i, row in enumerate(result_absence):
                national_id, name, rank, course, count = row
                absence_tree.insert("", tk.END, values=(i + 1, national_id, name, rank, course, count,
                                                        f"{(count / self.get_total_days(national_id) * 100):.1f}%"))

            # التأخير
            cursor.execute(query_lateness, course_params + [limit])
            result_lateness = cursor.fetchall()
            for i, row in enumerate(result_lateness):
                national_id, name, rank, course, count = row
                lateness_tree.insert("", tk.END, values=(i + 1, national_id, name, rank, course, count,
                                                         f"{(count / self.get_total_days(national_id) * 100):.1f}%"))

            # الغياب بعذر
            cursor.execute(query_excused, course_params + [limit])
            result_excused = cursor.fetchall()
            for i, row in enumerate(result_excused):
                national_id, name, rank, course, count = row
                excused_tree.insert("", tk.END, values=(i + 1, national_id, name, rank, course, count,
                                                        f"{(count / self.get_total_days(national_id) * 100):.1f}%"))

        # دالة تصدير البيانات
        def export_statistics():
            if not self.current_user["permissions"]["can_export_data"]:
                messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
                return

            limit = int(limit_var.get())
            course_filter = course_var.get()

            # اختيار مسار الملف
            export_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"إحصائيات_الغياب_والتأخير.xlsx"
            )

            if not export_file:
                return

            try:
                import pandas as pd
                from pandas import ExcelWriter

                # إعداد شرط الدورة للاستعلامات
                course_condition = ""
                course_params = []

                if course_filter != "جميع الدورات":
                    course_condition = "AND t.course = ?"
                    course_params = [course_filter]

                # استعلامات الحصول على البيانات
                query_absence = f"""
                SELECT t.national_id, t.name, t.rank, t.course, COUNT(*) as count
                FROM attendance a
                JOIN trainees t ON a.national_id = t.national_id
                WHERE a.status = 'غائب' AND t.is_excluded = 0 {course_condition}
                GROUP BY t.national_id
                ORDER BY count DESC
                LIMIT ?
                """

                query_lateness = f"""
                SELECT t.national_id, t.name, t.rank, t.course, COUNT(*) as count
                FROM attendance a
                JOIN trainees t ON a.national_id = t.national_id
                WHERE a.status = 'متأخر' AND t.is_excluded = 0 {course_condition}
                GROUP BY t.national_id
                ORDER BY count DESC
                LIMIT ?
                """

                query_excused = f"""
                SELECT t.national_id, t.name, t.rank, t.course, COUNT(*) as count
                FROM attendance a
                JOIN trainees t ON a.national_id = t.national_id
                WHERE a.status = 'غائب بعذر' AND t.is_excluded = 0 {course_condition}
                GROUP BY t.national_id
                ORDER BY count DESC
                LIMIT ?
                """

                cursor = self.conn.cursor()

                # استخراج البيانات من قاعدة البيانات
                cursor.execute(query_absence, course_params + [limit])
                absence_data = cursor.fetchall()

                cursor.execute(query_lateness, course_params + [limit])
                lateness_data = cursor.fetchall()

                cursor.execute(query_excused, course_params + [limit])
                excused_data = cursor.fetchall()

                # تحويل البيانات إلى DataFrame
                columns = ["رقم الهوية", "الاسم", "الرتبة", "الدورة", "عدد أيام الغياب", "النسبة المئوية"]

                # دالة مساعدة لإضافة النسبة المئوية
                def add_percentage(data_list):
                    result = []
                    for row in data_list:
                        national_id, name, rank, course, count = row
                        total_days = self.get_total_days(national_id)
                        percentage = f"{(count / total_days * 100):.1f}%" if total_days > 0 else "0%"
                        result.append([national_id, name, rank, course, count, percentage])
                    return result

                df_absence = pd.DataFrame(add_percentage(absence_data), columns=columns)
                df_lateness = pd.DataFrame(add_percentage(lateness_data), columns=columns)
                df_excused = pd.DataFrame(add_percentage(excused_data), columns=columns)

                # تصدير البيانات إلى Excel
                with ExcelWriter(export_file) as writer:
                    df_absence.to_excel(writer, sheet_name="أعلى معدلات الغياب", index=False)
                    df_lateness.to_excel(writer, sheet_name="أعلى معدلات التأخير", index=False)
                    df_excused.to_excel(writer, sheet_name="أعلى معدلات الغياب بعذر", index=False)

                messagebox.showinfo("نجاح", f"تم تصدير البيانات بنجاح إلى:\n{export_file}")

            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير البيانات: {str(e)}")

        # زر الإغلاق
        close_button = tk.Button(
            stats_window,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=20, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=stats_window.destroy
        )
        close_button.pack(pady=10)

        # تحميل البيانات مبدئياً
        load_statistics()

    def create_stats_tree(self, parent_frame):
        """إنشاء جدول عرض إحصائيات"""
        # إنشاء إطار للجدول مع شريط تمرير
        tree_frame = tk.Frame(parent_frame, bg=self.colors["light"])
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = tk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        tree = ttk.Treeview(
            tree_frame,
            columns=("rank", "id", "name", "grade", "course", "count", "percentage"),
            show="headings",
            yscrollcommand=scrollbar.set,
            style="Bold.Treeview"
        )

        # تعريف الأعمدة
        tree.column("rank", width=50, anchor=tk.CENTER)
        tree.column("id", width=120, anchor=tk.CENTER)
        tree.column("name", width=180, anchor=tk.CENTER)
        tree.column("grade", width=100, anchor=tk.CENTER)
        tree.column("course", width=150, anchor=tk.CENTER)
        tree.column("count", width=80, anchor=tk.CENTER)
        tree.column("percentage", width=100, anchor=tk.CENTER)

        # عناوين الأعمدة
        tree.heading("rank", text="الترتيب")
        tree.heading("id", text="رقم الهوية")
        tree.heading("name", text="الاسم")
        tree.heading("grade", text="الرتبة")
        tree.heading("course", text="الدورة")
        tree.heading("count", text="عدد المرات")
        tree.heading("percentage", text="النسبة")

        tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)

        return tree

    def get_total_days(self, national_id):
        """حساب إجمالي عدد أيام الدورة للطالب"""
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(DISTINCT date) 
            FROM attendance 
            WHERE national_id=?
        """, (national_id,))
        total = cursor.fetchone()[0]
        return total if total > 0 else 1  # لتجنب القسمة على صفر

    def export_based_on_filter(self):
        if not self.current_user["permissions"]["can_export_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
            return

        status_value = self.status_filter_var.get()
        if status_value == "الكل":
            self.export_filtered_records(None)
        else:
            self.export_filtered_records(status_value)

    def update_attendance_display(self):
        for row in self.attendance_tree.get_children():
            self.attendance_tree.delete(row)

        date_str = self.log_date_entry.get_date().strftime("%Y-%m-%d")
        df = pd.read_sql("""
            SELECT a.id, a.national_id, a.name, a.rank, a.course,
                   a.time, a.date, a.status, a.registered_by,
                   a.excuse_reason, a.updated_by, a.updated_at
            FROM attendance a
            JOIN trainees t ON a.national_id = t.national_id
            WHERE a.date=? AND t.is_excluded=0
        """, self.conn, params=[date_str])

        if df.empty:
            return

        status_filter_val = self.status_filter_var.get()
        if status_filter_val != "الكل":
            df = df[df['status'] == status_filter_val]

        search_text = self.log_search_var.get().strip()
        if search_text:
            df = df[df.apply(
                lambda row: (search_text in str(row['national_id'])) or (search_text in str(row['name'])),
                axis=1
            )]

        for _, row_ in df.iterrows():
            all_absences = self.get_all_absences_count(row_['national_id'])
            all_lates = self.get_all_late_count(row_['national_id'])
            all_excused = self.get_all_excused_count(row_['national_id'])

            if self.current_user["permissions"]["can_view_edit_history"]:
                values = (
                    row_['national_id'],
                    row_['name'],
                    row_['rank'],
                    row_['course'],
                    row_['status'],
                    all_absences,
                    all_lates,
                    all_excused,
                    row_['updated_by'] if row_['updated_by'] else "",
                    row_['updated_at'] if row_['updated_at'] else ""
                )
            else:
                values = (
                    row_['national_id'],
                    row_['name'],
                    row_['rank'],
                    row_['course'],
                    row_['status'],
                    all_absences,
                    all_lates,
                    all_excused,
                )

            item_id = self.attendance_tree.insert("", tk.END, values=values)
            st = row_['status']
            if st == "حاضر":
                self.attendance_tree.item(item_id, tags=("present",))
            elif st == "متأخر":
                self.attendance_tree.item(item_id, tags=("late",))
            elif st == "غائب بعذر":
                self.attendance_tree.item(item_id, tags=("excused",))
            elif st == "لم يباشر":
                self.attendance_tree.item(item_id, tags=("not_started",))
            elif st == "تطبيق ميداني":
                self.attendance_tree.item(item_id, tags=("field_application",))
            elif st == "يوم طالب":
                self.attendance_tree.item(item_id, tags=("student_day",))
            elif st == "مسائية / عن بعد":
                self.attendance_tree.item(item_id, tags=("evening_remote",))
            elif st == "حالة وفاة":
                self.attendance_tree.item(item_id, tags=("death_case",))
            elif st == "منوم":
                self.attendance_tree.item(item_id, tags=("hospital",))

    def reset_attendance(self):
        if not self.current_user["permissions"]["can_reset_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية إعادة تعيين سجلات الحضور")
            return

        choice = messagebox.askyesnocancel("إعادة تعيين السجل",
                                           "هل تريد إعادة تعيين السجل لليوم فقط؟\n\n(نعم = اليوم ، لا = كل الأيام ، إلغاء = لا شيء)")
        if choice is None:
            return
        if choice:
            date_str = self.log_date_entry.get_date().strftime("%Y-%m-%d")
            if messagebox.askyesno("تأكيد", f"سيتم حذف كل سجلات الحضور بتاريخ {date_str}، هل أنت متأكد؟"):
                try:
                    with self.conn:
                        self.conn.execute("DELETE FROM attendance WHERE date=?", (date_str,))
                    messagebox.showinfo("نجاح", f"تم حذف السجلات بتاريخ {date_str}")
                    self.update_attendance_display()
                    self.update_statistics()
                except Exception as e:
                    messagebox.showerror("خطأ", str(e))
        else:
            if messagebox.askyesno("تأكيد", "سيتم حذف كل سجلات الحضور لجميع الأيام، هل أنت متأكد؟"):
                try:
                    with self.conn:
                        self.conn.execute("DELETE FROM attendance")
                    messagebox.showinfo("نجاح", "تم حذف جميع سجلات الحضور")
                    self.update_attendance_display()
                    self.update_statistics()
                except Exception as e:
                    messagebox.showerror("خطأ", str(e))

    def export_filtered_records(self, status_value):
        if not self.current_user["permissions"]["can_export_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
            return

        date_str = self.log_date_entry.get_date().strftime("%Y-%m-%d")
        query = """
            SELECT
                a.national_id as 'رقم الهوية',
                a.name as 'الاسم',
                a.rank as 'الرتبة',
                a.course as 'اسم الدورة',
                a.time as 'وقت التسجيل',
                a.date as 'التاريخ',
                a.status as 'الحالة',
                a.registered_by as 'سجّل بواسطة',
                a.excuse_reason as 'السبب'
        """

        if self.current_user["permissions"]["can_view_edit_history"]:
            query += """,
                a.updated_by as 'آخر من عدّل',
                a.updated_at as 'وقت آخر تعديل'
            """

        query += """
            FROM attendance a
            JOIN trainees t ON a.national_id = t.national_id
            WHERE a.date=? AND t.is_excluded=0
        """

        params = [date_str]
        if status_value is not None:
            query += " AND a.status=?"
            params.append(status_value)

        df = pd.read_sql(query, self.conn, params=params)
        if df.empty:
            stxt = "الكل" if (status_value is None) else status_value
            messagebox.showinfo("ملاحظة", f"لا توجد بيانات {stxt} في هذا اليوم.")
            return

        stxt = "الكل" if (status_value is None) else status_value
        export_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{stxt}_{date_str}.xlsx"
        )
        if not export_file:
            return

        try:
            df.to_excel(export_file, index=False)
            messagebox.showinfo("نجاح", f"تم تصدير {stxt} إلى الملف: {export_file}")
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def on_attendance_double_click(self, event=None):
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تعديل سجلات الحضور")
            return

        item = self.attendance_tree.selection()
        if not item:
            return
        values = self.attendance_tree.item(item, "values")
        if not values:
            return

        national_id = values[0]
        date_str = self.log_date_entry.get_date().strftime("%Y-%m-%d")
        attendance_id = self.get_attendance_record_id(national_id, date_str)
        if not attendance_id:
            messagebox.showinfo("خطأ", "لا يمكن العثور على سجل الحضور المحدد.")
            return
        self.open_edit_attendance_window(attendance_id, values)

    def get_attendance_record_id(self, national_id, date_str):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM attendance WHERE national_id=? AND date=?", (national_id, date_str))
        result = cursor.fetchone()
        return result[0] if result else None

    def open_edit_attendance_window(self, attendance_id, row_values):
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تعديل سجلات الحضور")
            return

        # الحصول على الحالة الأصلية للطالب
        cursor = self.conn.cursor()
        cursor.execute("SELECT original_status FROM attendance WHERE id=?", (attendance_id,))
        original_status = cursor.fetchone()[0]

        edit_window = tk.Toplevel(self.root)
        edit_window.title("تعديل حالة الحضور")
        edit_window.geometry("500x430")  # زيادة الارتفاع لاستيعاب الحقل الجديد
        edit_window.configure(bg=self.colors["light"])
        edit_window.transient(self.root)
        edit_window.grab_set()

        x = (edit_window.winfo_screenwidth() - 500) // 2
        y = (edit_window.winfo_screenheight() - 430) // 2
        edit_window.geometry(f"500x430+{x}+{y}")

        tk.Label(edit_window, text="تعديل حالة الحضور", font=self.fonts["title"], bg=self.colors["primary"],
                 fg="white", padx=10, pady=10, width=500).pack(fill=tk.X)

        info_frame = tk.Frame(edit_window, bg=self.colors["light"], padx=20, pady=10)
        info_frame.pack(fill=tk.X)

        tk.Label(info_frame, text=f"الطالب: {row_values[1]}", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).pack(anchor=tk.W)
        tk.Label(info_frame, text=f"الحالة الحالية: {row_values[4]}", font=self.fonts["text"],
                 bg=self.colors["light"]).pack(anchor=tk.W)
        tk.Label(info_frame, text=f"الحالة الأصلية: {original_status}", font=self.fonts["text"],
                 bg=self.colors["light"]).pack(anchor=tk.W)

        new_status_frame = tk.LabelFrame(edit_window, text="اختر الحالة الجديدة", font=self.fonts["text_bold"],
                                         bg=self.colors["light"], fg=self.colors["dark"], padx=10, pady=10)
        new_status_frame.pack(fill=tk.X, padx=20, pady=10)

        status_options = ["حاضر", "متأخر", "غائب", "غائب بعذر", "لم يباشر"]
        status_var = tk.StringVar(value=row_values[4])

        status_combobox = ttk.Combobox(new_status_frame, textvariable=status_var, values=status_options,
                                       state="readonly",
                                       font=self.fonts["text"])
        status_combobox.pack(fill=tk.X, padx=5, pady=5)

        # إطار أسباب الغياب
        reason_frame = tk.Frame(edit_window, bg=self.colors["light"], padx=20, pady=5)
        reason_frame.pack(fill=tk.X)

        reason_label = tk.Label(reason_frame, text="سبب الغياب:", font=self.fonts["text"], bg=self.colors["light"])
        reason_entry = tk.Entry(reason_frame, font=self.fonts["text"], width=40)

        # إطار سبب التعديل
        mod_reason_frame = tk.Frame(edit_window, bg=self.colors["light"], padx=20, pady=5)
        mod_reason_frame.pack(fill=tk.X)

        mod_reason_label = tk.Label(mod_reason_frame, text="سبب التعديل:", font=self.fonts["text"],
                                    bg=self.colors["light"])
        mod_reason_entry = tk.Entry(mod_reason_frame, font=self.fonts["text"], width=40)

        def on_status_change(*args):
            # إظهار حقل سبب الغياب إذا كانت الحالة "غائب بعذر"
            if status_var.get() == "غائب بعذر":
                reason_label.pack(anchor=tk.W)
                reason_entry.pack(fill=tk.X, pady=5)
            else:
                reason_label.pack_forget()
                reason_entry.pack_forget()

            # إظهار حقل سبب التعديل فقط إذا كانت الحالة الجديدة مختلفة عن الحالة الأصلية
            if status_var.get() != original_status:
                mod_reason_label.pack(anchor=tk.W)
                mod_reason_entry.pack(fill=tk.X, pady=5)
            else:
                mod_reason_label.pack_forget()
                mod_reason_entry.pack_forget()

        status_var.trace("w", on_status_change)
        on_status_change()

        button_frame = tk.Frame(edit_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=20)

        def save_changes():
            new_status = status_var.get()
            new_reason = reason_entry.get().strip() if new_status == "غائب بعذر" else ""
            modification_reason = mod_reason_entry.get().strip() if new_status != original_status else ""
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                with self.conn:
                    self.conn.execute("""
                        UPDATE attendance
                        SET status=?, excuse_reason=?,
                            updated_by=?, updated_at=?, modification_reason=?
                        WHERE id=?
                    """, (
                        new_status,
                        new_reason,
                        self.current_user["full_name"],
                        now_str,
                        modification_reason,
                        attendance_id
                    ))
                messagebox.showinfo("نجاح", "تم تحديث حالة الحضور بنجاح")
                edit_window.destroy()
                self.update_attendance_display()
                self.update_statistics()
            except Exception as e:
                messagebox.showerror("خطأ", str(e))

        save_btn = tk.Button(button_frame, text="حفظ التغييرات", font=self.fonts["text_bold"],
                             bg=self.colors["success"],
                             fg="white", padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=save_changes)
        save_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = tk.Button(button_frame, text="إلغاء", font=self.fonts["text_bold"], bg=self.colors["danger"],
                               fg="white", padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2",
                               command=edit_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=5)

    def add_students_to_existing_course(self):
        """إضافة طلاب جدد لدورة موجودة مع خيار لتعيين الفصل مباشرة"""
        if not self.current_user["permissions"]["can_add_students"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية إضافة طلاب جدد")
            return

        # إنشاء نافذة جديدة
        add_window = tk.Toplevel(self.root)
        add_window.title("إضافة طلاب جدد لدورة موجودة")
        add_window.geometry("800x600")
        add_window.configure(bg=self.colors["light"])
        add_window.transient(self.root)
        add_window.grab_set()

        # توسيط النافذة
        x = (add_window.winfo_screenwidth() - 800) // 2
        y = (add_window.winfo_screenheight() - 600) // 2
        add_window.geometry(f"800x600+{x}+{y}")

        # إضافة عنوان للنافذة
        tk.Label(
            add_window,
            text="إضافة طلاب جدد لدورة موجودة",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        # إطار اختيار الدورة والفصل
        course_frame = tk.Frame(add_window, bg=self.colors["light"], padx=20, pady=10)
        course_frame.pack(fill=tk.X)

        tk.Label(course_frame, text="اختر الدورة:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=0,
                                                                                                                column=3,
                                                                                                                padx=5,
                                                                                                                pady=8,
                                                                                                                sticky=tk.E)

        # الحصول على قائمة الدورات المتاحة
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course FROM trainees")
        courses = [row[0] for row in cursor.fetchall() if row[0]]

        course_var = tk.StringVar()
        course_combo = ttk.Combobox(
            course_frame,
            textvariable=course_var,
            values=courses,
            font=self.fonts["text"],
            width=30,
            state="readonly"
        )
        course_combo.grid(row=0, column=2, padx=5, pady=8, sticky=tk.W)

        # إضافة عنصر اختيار الفصل
        tk.Label(course_frame, text="اختر الفصل:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=0,
                                                                                                               column=1,
                                                                                                               padx=5,
                                                                                                               pady=8,
                                                                                                               sticky=tk.E)

        section_var = tk.StringVar()
        section_combo = ttk.Combobox(
            course_frame,
            textvariable=section_var,
            font=self.fonts["text"],
            width=20,
            state="readonly"
        )
        section_combo.grid(row=0, column=0, padx=5, pady=8, sticky=tk.W)

        # دالة لتحديث قائمة الفصول عند اختيار دورة
        def update_sections():
            selected_course = course_var.get()
            if not selected_course:
                section_combo['values'] = []
                section_var.set("")
                return

            # الحصول على قائمة الفصول المتاحة للدورة
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT section_name FROM course_sections
                WHERE course_name=?
                ORDER BY section_name
            """, (selected_course,))
            sections = [row[0] for row in cursor.fetchall()]

            # لا نضيف خيار "بدون فصل" للقائمة المنسدلة
            section_combo['values'] = sections

            if sections:
                section_combo.current(0)  # اختيار أول فصل كافتراضي
            else:
                # إذا لم تكن هناك فصول، اطلب من المستخدم إنشاء فصل أولاً
                messagebox.showwarning("تنبيه",
                                       "لا توجد فصول لهذه الدورة. الرجاء إنشاء فصل أولاً من خلال 'إدارة الفصول وتصدير الكشوفات'.")
                return

        # ربط وظيفة تحديث الفصول بتغيير الدورة
        course_combo.bind("<<ComboboxSelected>>", lambda e: update_sections())

        # إضافة إطار معلومات الطلاب
        info_frame = tk.Frame(add_window, bg=self.colors["light"], padx=20, pady=5)
        info_frame.pack(fill=tk.X)

        tk.Label(
            info_frame,
            text="معلومات الطلاب المضافين:",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["primary"]
        ).pack(anchor=tk.W, pady=(0, 10))

        # إطار البحث
        search_frame = tk.Frame(add_window, bg=self.colors["light"])
        search_frame.pack(fill=tk.X, pady=5)

        tk.Label(search_frame, text="بحث بالاسم أو الهوية:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).pack(side=tk.RIGHT, padx=5)

        self.name_search_entry = tk.Entry(search_frame, font=self.fonts["text"], width=30, bd=2, relief=tk.GROOVE)
        self.name_search_entry.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        self.name_search_entry.bind("<KeyRelease>", self.dynamic_name_search)

        self.name_listbox = tk.Listbox(add_window, font=self.fonts["text"], height=4,
                                       selectbackground=self.colors["primary"], bd=2, relief=tk.GROOVE)
        self.name_listbox.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.name_listbox.bind("<<ListboxSelect>>", self.on_name_select)

        input_frame = tk.Frame(add_window, bg=self.colors["light"])
        input_frame.pack(fill=tk.X, pady=5)

        self.id_entry = tk.Entry(self.root)

        # تعديل: إضافة إطار للأزرار في الأسفل
        buttons_frame = tk.Frame(add_window, bg=self.colors["light"], pady=20)
        buttons_frame.pack(fill=tk.X, padx=20)

        add_one_btn = tk.Button(
            buttons_frame,
            text="إضافة طالب واحد",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: add_single_student(course_var.get(), section_var.get())
        )
        add_one_btn.pack(side=tk.RIGHT, padx=10)

        import_excel_btn = tk.Button(
            buttons_frame,
            text="استيراد من ملف Excel",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: import_from_excel(course_var.get(), section_var.get())
        )
        import_excel_btn.pack(side=tk.RIGHT, padx=10)

        close_btn = tk.Button(
            buttons_frame,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=add_window.destroy
        )
        close_btn.pack(side=tk.LEFT, padx=10)

        # دالة إضافة طالب واحد جديد للدورة
        def add_single_student(course_name, section_name):
            if not course_name:
                messagebox.showwarning("تنبيه", "الرجاء اختيار الدورة أولاً")
                return

            # استخدام نافذة إضافة طالب موجودة مع إعداد اسم الدورة مسبقاً
            single_window = tk.Toplevel(add_window)
            single_window.title("إضافة طالب جديد للدورة")
            single_window.geometry("400x350")
            single_window.configure(bg=self.colors["light"])
            single_window.transient(add_window)
            single_window.grab_set()

            x = (single_window.winfo_screenwidth() - 400) // 2
            y = (single_window.winfo_screenheight() - 350) // 2
            single_window.geometry(f"400x350+{x}+{y}")

            tk.Label(
                single_window,
                text=f"إضافة طالب جديد للدورة: {course_name}",
                font=self.fonts["title"],
                bg=self.colors["primary"],
                fg="white",
                padx=10, pady=10, width=400
            ).pack(fill=tk.X)

            form_frame = tk.Frame(single_window, bg=self.colors["light"], padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH)

            tk.Label(form_frame, text="رقم الهوية:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=0,
                                                                                                                 column=1,
                                                                                                                 padx=5,
                                                                                                                 pady=5,
                                                                                                                 sticky=tk.E)
            id_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
            id_entry.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

            tk.Label(form_frame, text="الاسم:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=1,
                                                                                                            column=1,
                                                                                                            padx=5,
                                                                                                            pady=5,
                                                                                                            sticky=tk.E)
            name_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
            name_entry.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)

            tk.Label(form_frame, text="الرتبة:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=2,
                                                                                                             column=1,
                                                                                                             padx=5,
                                                                                                             pady=5,
                                                                                                             sticky=tk.E)
            rank_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
            rank_entry.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)

            tk.Label(form_frame, text="رقم الجوال:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=3,
                                                                                                                 column=1,
                                                                                                                 padx=5,
                                                                                                                 pady=5,
                                                                                                                 sticky=tk.E)
            phone_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
            phone_entry.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

            button_frame = tk.Frame(single_window, bg=self.colors["light"], pady=10)
            button_frame.pack(fill=tk.X)

            def save_student():
                nid = id_entry.get().strip()
                name = name_entry.get().strip()
                rank_ = rank_entry.get().strip()
                phone = phone_entry.get().strip()
                selected_section = section_name  # الفصل المختار

                if not all([nid, name]):
                    messagebox.showwarning("تنبيه", "يجب إدخال رقم الهوية والاسم على الأقل")
                    return

                if not selected_section:
                    messagebox.showwarning("تنبيه", "يجب اختيار فصل للطالب")
                    return

                cursor = self.conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM trainees WHERE national_id=?", (nid,))
                exists = cursor.fetchone()[0]

                if exists > 0:
                    # التحقق مما إذا كان الطالب موجود في نفس الدورة
                    cursor.execute("SELECT course FROM trainees WHERE national_id=?", (nid,))
                    current_course = cursor.fetchone()[0]

                    if current_course == course_name:
                        messagebox.showwarning("تنبيه", f"رقم الهوية موجود بالفعل في نفس الدورة: {course_name}")
                        return

                    if not messagebox.askyesno("تأكيد الإضافة",
                                               f"الطالب برقم الهوية {nid} موجود في دورة أخرى: {current_course}\n\nهل تريد نقله من الدورة السابقة إلى الدورة الجديدة: {course_name}؟"):
                        return

                    try:
                        # حذف الطالب من الدورة القديمة
                        with self.conn:
                            # حذف سجلات الحضور للطالب
                            self.conn.execute("DELETE FROM attendance WHERE national_id=?", (nid,))
                            # حذف توزيع الفصول السابق
                            self.conn.execute("DELETE FROM student_sections WHERE national_id=?", (nid,))
                            # حذف الطالب نفسه
                            self.conn.execute("DELETE FROM trainees WHERE national_id=?", (nid,))
                    except Exception as e:
                        messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف السجل القديم: {str(e)}")
                        return

                current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                try:
                    with self.conn:
                        # إضافة الطالب للدورة
                        self.conn.execute("""
                            INSERT INTO trainees (national_id, name, rank, course, phone)
                            VALUES (?, ?, ?, ?, ?)
                        """, (nid, name, rank_, course_name, phone))

                        # إذا تم اختيار فصل (وليس "بدون فصل")، أضف الطالب إلى الفصل
                        if section_name and section_name != "بدون فصل":
                            self.conn.execute("""
                                INSERT INTO student_sections 
                                (national_id, course_name, section_name, assigned_date)
                                VALUES (?, ?, ?, ?)
                            """, (nid, course_name, section_name, current_date))

                    messagebox.showinfo("نجاح", f"تم إضافة الطالب {name} بنجاح" +
                                        (
                                            f" إلى فصل {section_name}" if section_name and section_name != "بدون فصل" else ""))
                    single_window.destroy()
                    self.update_students_tree()

                except Exception as e:
                    messagebox.showerror("خطأ", str(e))

            save_btn = tk.Button(button_frame, text="حفظ", font=self.fonts["text_bold"], bg=self.colors["success"],
                                 fg="white",
                                 padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=save_student)
            save_btn.pack(side=tk.LEFT, padx=10)

            cancel_btn = tk.Button(button_frame, text="إلغاء", font=self.fonts["text_bold"], bg=self.colors["danger"],
                                   fg="white",
                                   padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=single_window.destroy)
            cancel_btn.pack(side=tk.RIGHT, padx=10)

        def import_from_excel(course_name, section_name):
            """استيراد طلاب من ملف Excel مع إضافتهم إلى الفصل المحدد"""
            if not course_name:
                messagebox.showwarning("تنبيه", "الرجاء اختيار الدورة أولاً")
                return

            # اختيار ملف Excel
            file_path = filedialog.askopenfilename(
                title="اختر ملف Excel يحتوي على بيانات الطلاب",
                filetypes=[("Excel files", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("All files", "*.*")]
            )

            if not file_path:
                return

            # التحقق من وجود مكتبة pandas
            try:
                import pandas as pd
            except ImportError:
                messagebox.showerror("خطأ",
                                     "يجب تثبيت مكتبة pandas لاستيراد ملفات Excel. استخدم الأمر: pip install pandas openpyxl")
                return

            try:
                # قراءة الملف
                df = pd.read_excel(file_path)

                # التحقق من وجود الأعمدة المطلوبة
                required_columns = {"national_id", "name"}
                found_columns = set(df.columns)

                # دعم أسماء الأعمدة بالعربية
                arabic_columns = {
                    "رقم الهوية": "national_id",
                    "الاسم": "name",
                    "الرتبة": "rank",
                    "رقم الجوال": "phone"
                }

                # تغيير الأسماء العربية إلى إنجليزية إذا وجدت
                rename_dict = {}
                for arabic, english in arabic_columns.items():
                    if arabic in df.columns:
                        rename_dict[arabic] = english

                if rename_dict:
                    df = df.rename(columns=rename_dict)
                    found_columns = set(df.columns)

                missing_columns = required_columns - found_columns
                if missing_columns:
                    messagebox.showerror("خطأ",
                                         f"الأعمدة التالية مفقودة في الملف: {', '.join(missing_columns)}\n\nيجب أن يحتوي الملف على عمود 'national_id' (رقم الهوية) وعمود 'name' (الاسم) على الأقل.")
                    return

                # إنشاء نافذة تأكيد المعاينة
                preview_window = tk.Toplevel(add_window)
                preview_window.title("معاينة بيانات الطلاب")
                preview_window.geometry("800x600")
                preview_window.configure(bg=self.colors["light"])
                preview_window.transient(add_window)
                preview_window.grab_set()

                tk.Label(
                    preview_window,
                    text=f"معاينة الطلاب المراد إضافتهم إلى دورة: {course_name}" +
                         (f" - فصل: {section_name}" if section_name and section_name != "بدون فصل" else ""),
                    font=self.fonts["title"],
                    bg=self.colors["primary"],
                    fg="white",
                    padx=10, pady=10
                ).pack(fill=tk.X)

                # عرض إحصائيات حول البيانات
                total_records = len(df)
                tk.Label(
                    preview_window,
                    text=f"إجمالي الطلاب في الملف: {total_records}",
                    font=self.fonts["text_bold"],
                    bg=self.colors["light"],
                    pady=5
                ).pack()

                # إطار لعرض معاينة البيانات
                preview_frame = tk.Frame(preview_window, bg=self.colors["light"])
                preview_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

                preview_scroll = tk.Scrollbar(preview_frame)
                preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)

                # شجرة للمعاينة
                preview_tree = ttk.Treeview(
                    preview_frame,
                    columns=["id", "national_id", "name", "rank", "phone", "status"],
                    show="headings",
                    yscrollcommand=preview_scroll.set,
                    style="Bold.Treeview"
                )

                preview_tree.column("id", width=50, anchor=tk.CENTER)
                preview_tree.column("national_id", width=150, anchor=tk.CENTER)
                preview_tree.column("name", width=200, anchor=tk.CENTER)
                preview_tree.column("rank", width=100, anchor=tk.CENTER)
                preview_tree.column("phone", width=120, anchor=tk.CENTER)
                preview_tree.column("status", width=150, anchor=tk.CENTER)

                preview_tree.heading("id", text="#")
                preview_tree.heading("national_id", text="رقم الهوية")
                preview_tree.heading("name", text="الاسم")
                preview_tree.heading("rank", text="الرتبة")
                preview_tree.heading("phone", text="رقم الجوال")
                preview_tree.heading("status", text="الحالة")

                preview_tree.pack(fill=tk.BOTH, expand=True)
                preview_scroll.config(command=preview_tree.yview)

                # إضافة الإمكانية لتمييز الطلاب حسب حالتهم
                preview_tree.tag_configure("new", background="#e8f5e9")  # خلفية خضراء فاتحة للطلاب الجدد
                preview_tree.tag_configure("existing_same",
                                           background="#ffebee")  # خلفية حمراء فاتحة للطلاب الموجودين في نفس الدورة
                preview_tree.tag_configure("existing_other",
                                           background="#fff8e1")  # خلفية صفراء فاتحة للطلاب في دورات أخرى

                # العداد للطلاب حسب حالتهم
                new_students = 0
                existing_same_course = 0
                existing_other_course = 0

                # تحضير بيانات الطلاب للعرض
                student_rows = []

                for i, row in df.iterrows():
                    national_id = str(row["national_id"]).strip()
                    name = str(row["name"]).strip()

                    # إذا كان الصف يحتوي على بيانات فارغة أو غير صالحة، تخطيه
                    if not national_id or not name:
                        continue

                    rank = str(row.get("rank", "")).strip() if "rank" in row else ""
                    phone = str(row.get("phone", "")).strip() if "phone" in row else ""

                    # التحقق من وجود الطالب في قاعدة البيانات
                    cursor.execute("SELECT course FROM trainees WHERE national_id=?", (national_id,))
                    existing = cursor.fetchone()

                    status = ""
                    tag = ""

                    if existing:
                        existing_course = existing[0]
                        if existing_course == course_name:
                            status = f"موجود بالفعل في دورة: {existing_course}"
                            tag = "existing_same"
                            existing_same_course += 1
                        else:
                            status = f"موجود في دورة أخرى: {existing_course}"
                            tag = "existing_other"
                            existing_other_course += 1
                    else:
                        status = "جديد"
                        tag = "new"
                        new_students += 1

                    student_rows.append((i + 1, national_id, name, rank, phone, status, tag))

                # إضافة الصفوف إلى الشجرة
                for row in student_rows:
                    item_id = preview_tree.insert("", tk.END, values=row[:-1])
                    preview_tree.item(item_id, tags=(row[-1],))

                # إضافة ملخص الإحصائيات
                stats_frame = tk.Frame(preview_window, bg=self.colors["light"], padx=20, pady=10)
                stats_frame.pack(fill=tk.X)

                tk.Label(
                    stats_frame,
                    text=f"طلاب جدد: {new_students}",
                    font=self.fonts["text"],
                    bg="#e8f5e9", fg="black",
                    padx=10, pady=5
                ).pack(side=tk.RIGHT, padx=5)

                tk.Label(
                    stats_frame,
                    text=f"طلاب موجودون في نفس الدورة: {existing_same_course}",
                    font=self.fonts["text"],
                    bg="#ffebee", fg="black",
                    padx=10, pady=5
                ).pack(side=tk.RIGHT, padx=5)

                tk.Label(
                    stats_frame,
                    text=f"طلاب موجودون في دورات أخرى: {existing_other_course}",
                    font=self.fonts["text"],
                    bg="#fff8e1", fg="black",
                    padx=10, pady=5
                ).pack(side=tk.RIGHT, padx=5)

                # أزرار التأكيد أو الإلغاء
                button_frame = tk.Frame(preview_window, bg=self.colors["light"], pady=10)
                button_frame.pack(fill=tk.X, padx=10)

                # متغيرات الخيارات
                import_new_var = tk.IntVar(value=1)
                import_other_courses_var = tk.IntVar(value=0)
                import_same_course_var = tk.IntVar(value=0)

                # مربعات الاختيار
                tk.Checkbutton(
                    button_frame,
                    text="إضافة الطلاب الجدد",
                    variable=import_new_var,
                    font=self.fonts["text"],
                    bg=self.colors["light"]
                ).pack(anchor=tk.W)

                tk.Checkbutton(
                    button_frame,
                    text="إضافة الطلاب الموجودين في دورات أخرى (سيتم نقلهم)",
                    variable=import_other_courses_var,
                    font=self.fonts["text"],
                    bg=self.colors["light"]
                ).pack(anchor=tk.W)

                tk.Checkbutton(
                    button_frame,
                    text="تحديث بيانات الطلاب الموجودين في نفس الدورة",
                    variable=import_same_course_var,
                    font=self.fonts["text"],
                    bg=self.colors["light"]
                ).pack(anchor=tk.W)

                # إطار زر التنفيذ
                btn_frame = tk.Frame(preview_window, bg=self.colors["light"], pady=10)
                btn_frame.pack(fill=tk.X, padx=10)

                def execute_import():
                    # التحقق من تحديد خيار واحد على الأقل
                    if import_new_var.get() == 0 and import_other_courses_var.get() == 0 and import_same_course_var.get() == 0:
                        messagebox.showwarning("تنبيه", "الرجاء تحديد خيار واحد على الأقل للاستيراد")
                        return

                    # إنشاء نافذة تقدم العملية
                    progress_window = tk.Toplevel(preview_window)
                    progress_window.title("استيراد الطلاب")
                    progress_window.geometry("400x150")
                    progress_window.configure(bg=self.colors["light"])
                    progress_window.transient(preview_window)
                    progress_window.grab_set()

                    tk.Label(
                        progress_window,
                        text="جاري استيراد الطلاب...",
                        font=self.fonts["text_bold"],
                        bg=self.colors["light"],
                        pady=10
                    ).pack()

                    progress_var = tk.DoubleVar()
                    progress_bar = ttk.Progressbar(
                        progress_window,
                        variable=progress_var,
                        maximum=100,
                        length=350
                    )
                    progress_bar.pack(pady=10)

                    status_label = tk.Label(
                        progress_window,
                        text="جاري التحضير...",
                        font=self.fonts["text"],
                        bg=self.colors["light"]
                    )
                    status_label.pack(pady=5)

                    progress_window.update()

                    # حساب عدد العمليات المراد تنفيذها
                    operations_count = 0
                    if import_new_var.get() == 1:
                        operations_count += new_students
                    if import_other_courses_var.get() == 1:
                        operations_count += existing_other_course
                    if import_same_course_var.get() == 1:
                        operations_count += existing_same_course

                    current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    operations_done = 0
                    imported_new = 0
                    imported_from_other = 0
                    updated_same = 0
                    errors = 0

                    try:
                        with self.conn:
                            # معالجة الطلاب حسب نوعهم
                            for row_data in student_rows:
                                _, national_id, name, rank, phone, _, tag = row_data

                                # تحديث شريط التقدم
                                progress_var.set((operations_done / max(1, operations_count)) * 100)
                                status_label.config(text=f"معالجة الطالب: {name}")
                                progress_window.update()

                                # طلاب جدد
                                if tag == "new" and import_new_var.get() == 1:
                                    try:
                                        # إضافة الطالب للدورة
                                        self.conn.execute("""
                                            INSERT INTO trainees (national_id, name, rank, course, phone)
                                            VALUES (?, ?, ?, ?, ?)
                                        """, (national_id, name, rank, course_name, phone))

                                        # إذا تم اختيار فصل (غير "بدون فصل")، أضف الطالب للفصل
                                        if section_name and section_name != "بدون فصل":
                                            self.conn.execute("""
                                                INSERT INTO student_sections 
                                                (national_id, course_name, section_name, assigned_date)
                                                VALUES (?, ?, ?, ?)
                                            """, (national_id, course_name, section_name, current_date))

                                        imported_new += 1
                                    except Exception as e:
                                        print(f"خطأ في إضافة طالب جديد: {str(e)}")
                                        errors += 1

                                # طلاب من دورات أخرى
                                elif tag == "existing_other" and import_other_courses_var.get() == 1:
                                    try:
                                        # حذف سجلات الحضور للطالب
                                        self.conn.execute("DELETE FROM attendance WHERE national_id=?", (national_id,))
                                        # حذف توزيع الفصول السابق
                                        self.conn.execute("DELETE FROM student_sections WHERE national_id=?",
                                                          (national_id,))
                                        # تحديث معلومات الطالب مع الدورة الجديدة
                                        self.conn.execute("""
                                            UPDATE trainees
                                            SET name=?, rank=?, course=?, phone=?
                                            WHERE national_id=?
                                        """, (name, rank, course_name, phone, national_id))

                                        # إذا تم اختيار فصل (غير "بدون فصل")، أضف الطالب للفصل
                                        if section_name and section_name != "بدون فصل":
                                            self.conn.execute("""
                                                INSERT INTO student_sections 
                                                (national_id, course_name, section_name, assigned_date)
                                                VALUES (?, ?, ?, ?)
                                            """, (national_id, course_name, section_name, current_date))

                                        imported_from_other += 1
                                    except Exception as e:
                                        print(f"خطأ في نقل طالب من دورة أخرى: {str(e)}")
                                        errors += 1

                                # طلاب في نفس الدورة
                                elif tag == "existing_same" and import_same_course_var.get() == 1:
                                    try:
                                        # تحديث معلومات الطالب
                                        self.conn.execute("""
                                            UPDATE trainees
                                            SET name=?, rank=?, phone=?
                                            WHERE national_id=?
                                        """, (name, rank, phone, national_id))

                                        # التحقق مما إذا كان الطالب موجود في فصل حاليًا
                                        cursor.execute("""
                                            SELECT section_name FROM student_sections
                                            WHERE national_id=? AND course_name=?
                                        """, (national_id, course_name))
                                        current_section = cursor.fetchone()

                                        # إذا اختار المستخدم فصلًا غير "بدون فصل"
                                        if section_name and section_name != "بدون فصل":
                                            if current_section:
                                                # إذا كان الطالب في فصل مختلف، حدّث الفصل
                                                if current_section[0] != section_name:
                                                    self.conn.execute("""
                                                        UPDATE student_sections
                                                        SET section_name=?, assigned_date=?
                                                        WHERE national_id=? AND course_name=?
                                                    """, (section_name, current_date, national_id, course_name))
                                            else:
                                                # إذا لم يكن الطالب في أي فصل، أضفه إلى الفصل المحدد
                                                self.conn.execute("""
                                                    INSERT INTO student_sections 
                                                    (national_id, course_name, section_name, assigned_date)
                                                    VALUES (?, ?, ?, ?)
                                                """, (national_id, course_name, section_name, current_date))
                                        # إذا اختار المستخدم "بدون فصل" وكان الطالب في فصل
                                        elif section_name == "بدون فصل" and current_section:
                                            # إزالة الطالب من الفصل
                                            self.conn.execute("""
                                                DELETE FROM student_sections
                                                WHERE national_id=? AND course_name=?
                                            """, (national_id, course_name))

                                        updated_same += 1
                                    except Exception as e:
                                        print(f"خطأ في تحديث طالب موجود: {str(e)}")
                                        errors += 1

                                operations_done += 1

                        # إظهار ملخص النتائج
                        progress_var.set(100)
                        status_label.config(text="تم استيراد الطلاب بنجاح!")
                        progress_window.update()

                        # إغلاق نافذة التقدم بعد ثانيتين
                        progress_window.after(2000, progress_window.destroy)

                        # عرض ملخص النتائج
                        summary = f"تم إكمال عملية الاستيراد بنجاح:\n\n"
                        if import_new_var.get() == 1:
                            summary += f"• تم إضافة {imported_new} طالب جديد\n"
                        if import_other_courses_var.get() == 1:
                            summary += f"• تم نقل {imported_from_other} طالب من دورات أخرى\n"
                        if import_same_course_var.get() == 1:
                            summary += f"• تم تحديث بيانات {updated_same} طالب موجود\n"
                        if errors > 0:
                            summary += f"\nملاحظة: حدث {errors} أخطاء أثناء الاستيراد"

                        if section_name and section_name != "بدون فصل":
                            summary += f"\n\nتم توزيع الطلاب على فصل: {section_name}"

                        messagebox.showinfo("نتائج الاستيراد", summary)

                        # تحديث عرض الطلاب
                        self.update_students_tree()

                        # إغلاق نافذة المعاينة
                        preview_window.destroy()

                    except Exception as e:
                        # في حالة حدوث خطأ
                        try:
                            progress_window.destroy()
                        except:
                            pass
                        messagebox.showerror("خطأ", f"حدث خطأ أثناء الاستيراد: {str(e)}")

                confirm_btn = tk.Button(
                    btn_frame,
                    text="تنفيذ الاستيراد",
                    font=self.fonts["text_bold"],
                    bg=self.colors["success"],
                    fg="white",
                    padx=15, pady=5,
                    bd=0, relief=tk.FLAT,
                    cursor="hand2",
                    command=execute_import
                )
                confirm_btn.pack(side=tk.LEFT, padx=5)

                cancel_btn = tk.Button(
                    btn_frame,
                    text="إلغاء",
                    font=self.fonts["text_bold"],
                    bg=self.colors["danger"],
                    fg="white",
                    padx=15, pady=5,
                    bd=0, relief=tk.FLAT,
                    cursor="hand2",
                    command=preview_window.destroy
                )
                cancel_btn.pack(side=tk.RIGHT, padx=5)

            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء قراءة ملف Excel: {str(e)}")

        # تحديث قائمة الفصول عند فتح النافذة
        update_sections()

    def close_window(self):
        # إلغاء ربط الحدث قبل إغلاق النافذة
        try:
            self.name_search_entry.unbind("<KeyRelease>")
        except (tk.TclError, AttributeError):
            pass
        # إغلاق النافذة
        self.window.destroy()

    def is_widget_valid(self, widget):
        """التحقق من صلاحية عنصر واجهة المستخدم"""
        try:
            widget.winfo_exists()
            return True
        except (tk.TclError, AttributeError):
            return False

    def setup_students_tab(self):
        search_frame = tk.LabelFrame(self.students_tab, text="البحث عن طالب", font=self.fonts["subtitle"],
                                     bg=self.colors["light"], fg=self.colors["dark"], padx=10, pady=10)
        search_frame.pack(fill=tk.X, padx=10, pady=5)

        search_inner_frame = tk.Frame(search_frame, bg=self.colors["light"])
        search_inner_frame.pack(fill=tk.X, padx=5, pady=5)

        # تعديل: تغيير القائمة المنسدلة لعرض جميع الطلاب والطلاب المستبعدين
        tk.Label(search_inner_frame, text="عرض:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).pack(side=tk.RIGHT, padx=5)

        self.course_type_var = tk.StringVar(value="جميع الطلاب")
        course_type_combo = ttk.Combobox(
            search_inner_frame,
            textvariable=self.course_type_var,
            values=["جميع الطلاب", "الطلاب النشطين", "الطلاب المستبعدين"],  # تغيير الخيارات
            state="readonly",
            width=20,
            font=self.fonts["text"]
        )
        course_type_combo.pack(side=tk.RIGHT, padx=5)
        course_type_combo.bind("<<ComboboxSelected>>", lambda e: self.update_students_tree())

        tk.Label(search_inner_frame, text="بحث (الاسم أو الهوية):", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).pack(side=tk.RIGHT, padx=5)
        self.search_entry = tk.Entry(search_inner_frame, font=self.fonts["text"], width=30, bd=2, relief=tk.GROOVE)
        self.search_entry.pack(side=tk.RIGHT, padx=5)
        self.search_entry.bind("<Return>", lambda e: self.search_student())

        search_button = tk.Button(
            search_inner_frame, text="بحث", font=self.fonts["text_bold"], bg=self.colors["primary"], fg="white",
            padx=10, pady=3, bd=0, relief=tk.FLAT, command=self.search_student, cursor="hand2"
        )
        search_button.pack(side=tk.RIGHT, padx=5)

        show_all_button = tk.Button(
            search_inner_frame, text="عرض الكل", font=self.fonts["text_bold"], bg=self.colors["secondary"], fg="white",
            padx=10, pady=3, bd=0, relief=tk.FLAT, command=self.update_students_tree, cursor="hand2"
        )
        show_all_button.pack(side=tk.LEFT, padx=5)

        button_frame = tk.Frame(search_frame, bg=self.colors["light"])
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        add_to_course_btn = tk.Button(
            button_frame,
            text="إضافة طلاب جدد لدورة موجودة",
            font=self.fonts["text_bold"],
            bg="#3949AB",  # لون مميز
            fg="white",
            padx=10, pady=3,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=self.add_students_to_existing_course
        )
        add_to_course_btn.pack(side=tk.LEFT, padx=5)

        if self.current_user["permissions"]["can_add_students"]:
            add_button = tk.Button(
                button_frame, text="إضافة طالب جديد", font=self.fonts["text_bold"], bg=self.colors["success"],
                fg="white",
                padx=10, pady=3, bd=0, relief=tk.FLAT, command=self.add_new_student, cursor="hand2"
            )
            add_button.pack(side=tk.RIGHT, padx=5)

        if self.current_user["permissions"]["can_edit_students"]:
            edit_button = tk.Button(
                button_frame, text="تعديل الطالب المحدد", font=self.fonts["text_bold"], bg=self.colors["late"],
                fg="white",
                padx=10, pady=3, bd=0, relief=tk.FLAT, command=lambda: self.edit_student(from_selection=True),
                cursor="hand2"
            )
            edit_button.pack(side=tk.RIGHT, padx=5)

        if self.current_user["permissions"]["can_delete_students"]:
            delete_button = tk.Button(
                button_frame, text="حذف الطالب المحدد", font=self.fonts["text_bold"], bg=self.colors["danger"],
                fg="white",
                padx=10, pady=3, bd=0, relief=tk.FLAT, command=self.delete_selected_student, cursor="hand2"
            )
            delete_button.pack(side=tk.RIGHT, padx=5)

            multi_courses_button = tk.Button(
                button_frame, text="إدارة الفصول وتصدير الكشوفات", font=self.fonts["text_bold"], bg="#4285f4",
                # لون أزرق أكثر بروزًا
                fg="white",
                padx=10, pady=3, bd=0, relief=tk.FLAT, cursor="hand2", command=self.manage_multi_section_courses
            )
            multi_courses_button.pack(side=tk.LEFT, padx=5)

        if self.current_user["permissions"]["can_import_data"]:
            import_course_button = tk.Button(
                button_frame, text="استيراد دورة جديدة", font=self.fonts["text_bold"], bg=self.colors["secondary"],
                fg="white",
                padx=10, pady=3, bd=0, relief=tk.FLAT, command=self.import_new_course, cursor="hand2"
            )
            import_course_button.pack(side=tk.LEFT, padx=5)

        view_profile_button = tk.Button(
            button_frame, text="عرض ملف الطالب", font=self.fonts["text_bold"], bg=self.colors["secondary"], fg="white",
            padx=10, pady=3, bd=0, relief=tk.FLAT, command=self.view_student_profile, cursor="hand2"
        )
        view_profile_button.pack(side=tk.RIGHT, padx=5)

        students_display_frame = tk.LabelFrame(self.students_tab, text="قائمة الطلاب", font=self.fonts["subtitle"],
                                               bg=self.colors["light"], fg=self.colors["dark"], padx=10, pady=10)
        students_display_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.students_tree_scroll = tk.Scrollbar(students_display_frame)
        self.students_tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # إضافة عمود لإظهار حالة الاستبعاد
        self.students_tree = ttk.Treeview(
            students_display_frame,
            columns=("id", "name", "rank", "course", "phone", "status"),
            show="headings",
            yscrollcommand=self.students_tree_scroll.set,
            style="Bold.Treeview"
        )
        self.students_tree.column("id", width=120, anchor=tk.CENTER)
        self.students_tree.column("name", width=150, anchor=tk.CENTER)
        self.students_tree.column("rank", width=80, anchor=tk.CENTER)
        self.students_tree.column("course", width=150, anchor=tk.CENTER)
        self.students_tree.column("phone", width=120, anchor=tk.CENTER)
        self.students_tree.column("status", width=80, anchor=tk.CENTER)

        self.students_tree.heading("id", text="رقم الهوية")
        self.students_tree.heading("name", text="الاسم")
        self.students_tree.heading("rank", text="الرتبة")
        self.students_tree.heading("course", text="اسم الدورة")
        self.students_tree.heading("phone", text="رقم الجوال")
        self.students_tree.heading("status", text="الحالة")

        self.students_tree.pack(fill=tk.BOTH, expand=True)
        self.students_tree_scroll.config(command=self.students_tree.yview)

        # إضافة نمط للطلاب المستبعدين
        self.students_tree.tag_configure("excluded", background="#f8bbd0")

        self.students_tree.bind("<Double-1>", self.on_student_double_click)

    def setup_archive_tab(self):
        """إعداد تبويب أرشيف الدورات"""
        archive_frame = tk.Frame(self.archive_tab, bg=self.colors["light"], padx=10, pady=10)
        archive_frame.pack(fill=tk.BOTH, expand=True)

        # عنوان التبويب
        tk.Label(
            archive_frame,
            text="أرشيف الدورات التدريبية",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)
        # إطار أزرار الأرشفة
        button_frame = tk.Frame(archive_frame, bg=self.colors["light"], pady=20)
        button_frame.pack(fill=tk.X)

        # قسم تصدير الدورات إلى الأرشيف
        export_labelframe = tk.LabelFrame(
            archive_frame,
            text="تصدير الدورات إلى الأرشيف",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=15,
            pady=15
        )
        export_labelframe.pack(fill=tk.X, pady=10)

        tk.Label(
            export_labelframe,
            text="اختر الدورة التي تريد تصديرها إلى أرشيف خارجي:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(anchor=tk.W, pady=(0, 10))

        # إطار لعرض الدورات المتاحة
        courses_frame = tk.Frame(export_labelframe, bg=self.colors["light"])
        courses_frame.pack(fill=tk.X, pady=5)

        # الحصول على جميع الدورات الحالية
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course FROM trainees")
        all_courses = [row[0] for row in cursor.fetchall() if row[0]]

        # إطار القائمة
        list_frame = tk.Frame(export_labelframe, bg=self.colors["light"])
        list_frame.pack(fill=tk.X, pady=5)

        # شريط تمرير للقائمة
        list_scroll = tk.Scrollbar(list_frame)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # قائمة الدورات مع إمكانية اختيار متعدد
        self.archive_courses_listbox = tk.Listbox(
            list_frame,
            font=self.fonts["text"],
            selectbackground=self.colors["primary"],
            selectforeground="white",
            selectmode=tk.MULTIPLE,
            height=6,
            yscrollcommand=list_scroll.set
        )
        self.archive_courses_listbox.pack(fill=tk.X)
        list_scroll.config(command=self.archive_courses_listbox.yview)

        # إضافة الدورات إلى القائمة
        for course in all_courses:
            self.archive_courses_listbox.insert(tk.END, course)

        # أزرار التصدير
        buttons_frame = tk.Frame(export_labelframe, bg=self.colors["light"], pady=10)
        buttons_frame.pack(fill=tk.X)

        # دالة تصدير الدورات المحددة
        def export_selected_courses():
            selected_indices = self.archive_courses_listbox.curselection()
            if not selected_indices:
                messagebox.showinfo("تنبيه", "الرجاء اختيار دورة واحدة على الأقل للتصدير")
                return

            selected_courses = [self.archive_courses_listbox.get(idx) for idx in selected_indices]

            # تأكيد التصدير
            confirm_msg = f"هل أنت متأكد من تصدير الدورات التالية إلى الأرشيف؟\n\n"
            for course in selected_courses:
                confirm_msg += f"- {course}\n"

            if not messagebox.askyesno("تأكيد التصدير", confirm_msg):
                return

            # تصدير الدورات المحددة
            success = self.archive_manager.export_courses_to_archive(selected_courses)

            if success:
                # سؤال المستخدم إذا كان يريد حذف الدورات من النظام
                if messagebox.askyesno(
                        "حذف الدورات",
                        "تم تصدير الدورات بنجاح إلى الأرشيف. هل تريد حذف هذه الدورات من النظام الآن؟"
                ):
                    self.delete_courses_after_archive(selected_courses)

                # تحديث قائمة الدورات
                self.update_archive_courses_list()

        # أزرار الأرشفة
        export_selected_btn = tk.Button(
            buttons_frame,
            text="تصدير الدورات المحددة",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=export_selected_courses
        )
        export_selected_btn.pack(side=tk.LEFT, padx=5)

        refresh_btn = tk.Button(
            buttons_frame,
            text="تحديث القائمة",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=self.update_archive_courses_list
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)

        # قسم استعراض الأرشيف
        view_labelframe = tk.LabelFrame(
            archive_frame,
            text="استعراض الأرشيف",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=15,
            pady=15
        )
        view_labelframe.pack(fill=tk.X, pady=10)

        tk.Label(
            view_labelframe,
            text="استعراض الدورات المؤرشفة سابقًا (قراءة فقط):",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(anchor=tk.W, pady=(0, 10))

        view_buttons_frame = tk.Frame(view_labelframe, bg=self.colors["light"])
        view_buttons_frame.pack(fill=tk.X, pady=5)

        open_archive_btn = tk.Button(
            view_buttons_frame,
            text="فتح ملف أرشيف دورات",
            font=self.fonts["text_bold"],
            bg="#FF5722",  # لون برتقالي للتمييز
            fg="white",
            padx=15, pady=8,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=self.archive_manager.open_archive_window
        )
        open_archive_btn.pack(side=tk.RIGHT, padx=5)

        # إضافة زر دمج ملفات الأرشيف
        merge_archives_btn = tk.Button(
            view_buttons_frame,
            text="دمج ملفات أرشيف",
            font=self.fonts["text_bold"],
            bg="#9C27B0",  # لون بنفسجي للتمييز
            fg="white",
            padx=15, pady=8,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=self.archive_manager.merge_archives
        )
        merge_archives_btn.pack(side=tk.RIGHT, padx=5)

    def update_archive_courses_list(self):
        """تحديث قائمة الدورات المتاحة للأرشفة"""
        self.archive_courses_listbox.delete(0, tk.END)

        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course FROM trainees")
        all_courses = [row[0] for row in cursor.fetchall() if row[0]]

        for course in all_courses:
            self.archive_courses_listbox.insert(tk.END, course)

    def delete_courses_after_archive(self, course_names):
        """حذف الدورات بعد أرشفتها"""
        if not self.current_user["permissions"]["can_delete_students"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية حذف الدورات")
            return

        if not course_names:
            return

        # إنشاء نافذة تقدم العملية
        progress_window = tk.Toplevel(self.root)
        progress_window.title("حذف الدورات")
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
            text=f"جاري حذف {len(course_names)} دورة...",
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
            text="جاري تحضير العملية...",
            font=self.fonts["text"],
            bg=self.colors["light"]
        )
        status_label.pack(pady=5)

        progress_window.update()

        try:
            cursor = self.conn.cursor()

            # الحذف لكل دورة
            for i, course_name in enumerate(course_names):
                # تحديث شريط التقدم
                progress_var.set((i / len(course_names)) * 80)
                status_label.config(text=f"جاري حذف دورة: {course_name}")
                progress_window.update()

                # الحصول على أرقام هويات الطلاب في الدورة
                cursor.execute("SELECT national_id FROM trainees WHERE course=?", (course_name,))
                student_ids = [row[0] for row in cursor.fetchall()]

                with self.conn:
                    # حذف سجلات الحضور للطلاب
                    for student_id in student_ids:
                        self.conn.execute("DELETE FROM attendance WHERE national_id=?", (student_id,))

                    # حذف توزيع الطلاب على الفصول
                    self.conn.execute("DELETE FROM student_sections WHERE course_name=?", (course_name,))

                    # حذف الفصول
                    self.conn.execute("DELETE FROM course_sections WHERE course_name=?", (course_name,))

                    # حذف الطلاب
                    self.conn.execute("DELETE FROM trainees WHERE course=?", (course_name,))

            progress_var.set(100)
            status_label.config(text="تم حذف الدورات بنجاح!")
            progress_window.update()

            # إغلاق نافذة التقدم بعد ثانيتين
            progress_window.after(2000, progress_window.destroy)

            # تحديث البيانات
            self.update_students_tree()
            self.update_statistics()
            self.update_attendance_display()

            messagebox.showinfo("نجاح", f"تم حذف {len(course_names)} دورة بنجاح")

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف الدورات: {str(e)}")
            progress_window.destroy()

    def on_student_double_click(self, event):
        selected_item = self.students_tree.selection()
        if not selected_item:
            return
        self.view_student_profile()

    def search_student(self):
        """البحث عن طالب مع اعتبار فلتر نوع القائمة الجديدة"""
        text = self.search_entry.get().strip()
        filter_type = self.course_type_var.get()

        # تعديل: شرط البحث حسب الخيار المحدد
        if filter_type == "الطلاب النشطين":
            where_clause = "WHERE (national_id LIKE ? OR name LIKE ?) AND is_excluded=0"
        elif filter_type == "الطلاب المستبعدين":
            where_clause = "WHERE (national_id LIKE ? OR name LIKE ?) AND is_excluded=1"
        else:  # "جميع الطلاب"
            where_clause = "WHERE (national_id LIKE ? OR name LIKE ?)"

        query = f"""
        SELECT t.*, 
            (SELECT section_name FROM student_sections 
             WHERE national_id = t.national_id AND course_name = t.course) as section_name
        FROM trainees t
        {where_clause}
        """

        df = pd.read_sql(query, self.conn, params=[f"%{text}%", f"%{text}%"])
        self.populate_students_tree(df)

    def update_students_tree(self):
        """تحديث شجرة الطلاب مع خيارات التصفية الجديدة"""
        filter_type = self.course_type_var.get()

        # تعديل: التصفية حسب الخيار المحدد (الكل أو النشطين أو المستبعدين)
        if filter_type == "الطلاب النشطين":
            query = """
            SELECT t.*, 
                (SELECT section_name FROM student_sections 
                 WHERE national_id = t.national_id AND course_name = t.course) as section_name
            FROM trainees t
            WHERE t.is_excluded=0
            """
        elif filter_type == "الطلاب المستبعدين":
            query = """
            SELECT t.*, 
                (SELECT section_name FROM student_sections 
                 WHERE national_id = t.national_id AND course_name = t.course) as section_name
            FROM trainees t
            WHERE t.is_excluded=1
            """
        else:  # "جميع الطلاب"
            query = """
            SELECT t.*, 
                (SELECT section_name FROM student_sections 
                 WHERE national_id = t.national_id AND course_name = t.course) as section_name
            FROM trainees t
            """

        df = pd.read_sql(query, self.conn)
        self.populate_students_tree(df)

    def populate_students_tree(self, df):
        """تعبئة شجرة الطلاب مع إضافة معلومات الفصل للطلاب في الدورات متعددة الفصول"""
        for item in self.students_tree.get_children():
            self.students_tree.delete(item)

        for _, row in df.iterrows():
            # تحديد حالة الطالب (عادي أو مستبعد)
            status = "مستبعد" if row['is_excluded'] == 1 else "موجود"

            # تحضير عرض اسم الدورة مع الفصل إذا كان متوفراً
            course_display = row['course']
            if pd.notna(row.get('section_name')):
                course_display = f"{course_display} - فصل: {row['section_name']}"

            values = (row['national_id'], row['name'], row['rank'], course_display, row['phone'], status)
            item_id = self.students_tree.insert("", tk.END, values=values)

            # تمييز الطلاب المستبعدين
            if row['is_excluded'] == 1:
                self.students_tree.item(item_id, tags=("excluded",))

    def add_new_student(self):
        if not self.current_user["permissions"]["can_add_students"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية إضافة طلاب جدد")
            return

        add_window = tk.Toplevel(self.root)
        add_window.title("إضافة طالب جديد")
        add_window.geometry("400x350")
        add_window.configure(bg=self.colors["light"])
        add_window.transient(self.root)
        add_window.grab_set()

        x = (add_window.winfo_screenwidth() - 400) // 2
        y = (add_window.winfo_screenheight() - 350) // 2
        add_window.geometry(f"400x350+{x}+{y}")

        tk.Label(
            add_window,
            text="إضافة طالب جديد",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10, width=400
        ).pack(fill=tk.X)

        form_frame = tk.Frame(add_window, bg=self.colors["light"], padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH)

        tk.Label(form_frame, text="رقم الهوية:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=0,
                                                                                                             column=1,
                                                                                                             padx=5,
                                                                                                             pady=5,
                                                                                                             sticky=tk.E)
        id_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        id_entry.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="الاسم:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=1, column=1,
                                                                                                        padx=5, pady=5,
                                                                                                        sticky=tk.E)
        name_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        name_entry.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="الرتبة:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=2,
                                                                                                         column=1,
                                                                                                         padx=5, pady=5,
                                                                                                         sticky=tk.E)
        rank_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        rank_entry.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="اسم الدورة:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=3,
                                                                                                             column=1,
                                                                                                             padx=5,
                                                                                                             pady=5,
                                                                                                             sticky=tk.E)

        # عرض قائمة الدورات الحالية
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course FROM trainees")
        courses = [row[0] for row in cursor.fetchall() if row[0]]

        if courses:
            course_entry = ttk.Combobox(form_frame, font=self.fonts["text"], width=23, values=courses)
        else:
            course_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)

        course_entry.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="رقم الجوال:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=4,
                                                                                                             column=1,
                                                                                                             padx=5,
                                                                                                             pady=5,
                                                                                                             sticky=tk.E)
        phone_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        phone_entry.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)

        button_frame = tk.Frame(add_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X)

        def save_student():
            nid = id_entry.get().strip()
            name = name_entry.get().strip()
            rank_ = rank_entry.get().strip()
            course = course_entry.get().strip()
            phone = phone_entry.get().strip()
            if not nid or not name:
                messagebox.showwarning("تنبيه", "يجب إدخال رقم الهوية والاسم على الأقل")
                return
            cursor = self.conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM trainees WHERE national_id=?", (nid,))
            exists = cursor.fetchone()[0]

            # في حالة وجود الطالب بالفعل، نسأل المستخدم إذا كان يريد الاستمرار
            if exists > 0:
                # تحقق ما إذا كان الطالب موجود بنفس اسم الدورة
                cursor.execute("SELECT course FROM trainees WHERE national_id=?", (nid,))
                current_course = cursor.fetchone()[0]

                if current_course == course:
                    messagebox.showwarning("تنبيه", f"رقم الهوية موجود بالفعل في نفس الدورة: {course}")
                    return

                if not messagebox.askyesno("تأكيد الإضافة",
                                           f"الطالب برقم الهوية {nid} موجود في دورة أخرى: {current_course}\n\nهل تريد حذفه من الدورة السابقة وإضافته للدورة الجديدة: {course}؟"):
                    return

                try:
                    # حذف الطالب من الدورة القديمة
                    with self.conn:
                        # حذف سجلات الحضور للطالب
                        self.conn.execute("DELETE FROM attendance WHERE national_id=?", (nid,))
                        # حذف الطالب نفسه
                        self.conn.execute("DELETE FROM trainees WHERE national_id=?", (nid,))
                except Exception as e:
                    messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف السجل القديم: {str(e)}")
                    return

            try:
                with self.conn:
                    self.conn.execute("""
                        INSERT INTO trainees (national_id, name, rank, course, phone)
                        VALUES (?, ?, ?, ?, ?)
                    """, (nid, name, rank_, course, phone))
                messagebox.showinfo("نجاح", "تمت الإضافة بنجاح")
                add_window.destroy()
                self.update_students_tree()
                self.update_statistics()
            except Exception as e:
                messagebox.showerror("خطأ", str(e))

        save_btn = tk.Button(button_frame, text="حفظ", font=self.fonts["text_bold"], bg=self.colors["success"],
                             fg="white",
                             padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=save_student)
        save_btn.pack(side=tk.LEFT, padx=10)

        cancel_btn = tk.Button(button_frame, text="إلغاء", font=self.fonts["text_bold"], bg=self.colors["danger"],
                               fg="white",
                               padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=add_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=10)

    def edit_student(self, from_selection=False):
        if not self.current_user["permissions"]["can_edit_students"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تعديل بيانات الطلاب")
            return

        if from_selection:
            selected_item = self.students_tree.selection()
            if not selected_item:
                messagebox.showinfo("تنبيه", "الرجاء تحديد طالب من القائمة")
                return
            values = self.students_tree.item(selected_item, "values")
            nid = values[0]
        else:
            nid = simpledialog.askstring("تعديل طالب", "أدخل رقم هوية الطالب:")
            if not nid:
                return

        cursor = self.conn.cursor()

        # أولاً: استرجاع معلومات الطالب
        cursor.execute("SELECT * FROM trainees WHERE national_id=?", (nid,))
        student = cursor.fetchone()

        if not student:
            messagebox.showinfo("تنبيه", "لا توجد معلومات عن هذا الطالب")
            return

        # ثانياً: استرجاع سجلات الحضور بشكل منفصل ومباشر
        cursor.execute("""
                SELECT id, national_id, name, rank, course, time, date, status, original_status, 
                       registered_by, excuse_reason, updated_by, updated_at, modification_reason
                FROM attendance 
                WHERE national_id=?
                ORDER BY date DESC
            """, (nid,))
        attendance_records = cursor.fetchall()

        edit_window = tk.Toplevel(self.root)
        edit_window.title("تعديل بيانات الطالب")
        edit_window.geometry("400x350")
        edit_window.configure(bg=self.colors["light"])
        edit_window.transient(self.root)
        edit_window.grab_set()

        x = (edit_window.winfo_screenwidth() - 400) // 2
        y = (edit_window.winfo_screenheight() - 350) // 2
        edit_window.geometry(f"400x350+{x}+{y}")

        tk.Label(
            edit_window,
            text="تعديل بيانات الطالب",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10, width=400
        ).pack(fill=tk.X)

        form_frame = tk.Frame(edit_window, bg=self.colors["light"], padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH)

        tk.Label(form_frame, text="رقم الهوية:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=0,
                                                                                                             column=1,
                                                                                                             padx=5,
                                                                                                             pady=5,
                                                                                                             sticky=tk.E)
        id_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        id_entry.insert(0, student[0])
        id_entry.config(state="disabled")
        id_entry.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="الاسم:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=1, column=1,
                                                                                                        padx=5, pady=5,
                                                                                                        sticky=tk.E)
        name_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        name_entry.insert(0, student[1])
        name_entry.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="الرتبة:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=2,
                                                                                                         column=1,
                                                                                                         padx=5, pady=5,
                                                                                                         sticky=tk.E)
        rank_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        rank_entry.insert(0, student[2])
        rank_entry.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="اسم الدورة:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=3,
                                                                                                             column=1,
                                                                                                             padx=5,
                                                                                                             pady=5,
                                                                                                             sticky=tk.E)

        # عرض قائمة الدورات الحالية
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course FROM trainees")
        courses = [row[0] for row in cursor.fetchall() if row[0]]

        if courses:
            course_entry = ttk.Combobox(form_frame, font=self.fonts["text"], width=23, values=courses)
            course_entry.set(student[3])
        else:
            course_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
            course_entry.insert(0, student[3])

        course_entry.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

        tk.Label(form_frame, text="رقم الجوال:", font=self.fonts["text_bold"], bg=self.colors["light"]).grid(row=4,
                                                                                                             column=1,
                                                                                                             padx=5,
                                                                                                             pady=5,
                                                                                                             sticky=tk.E)
        phone_entry = tk.Entry(form_frame, font=self.fonts["text"], width=25)
        phone_entry.insert(0, student[4])
        phone_entry.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)

        button_frame = tk.Frame(edit_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X)

        def save_edit():
            new_name = name_entry.get().strip()
            new_rank = rank_entry.get().strip()
            new_course = course_entry.get().strip()
            new_phone = phone_entry.get().strip()
            if not new_name:
                messagebox.showwarning("تنبيه", "لا يمكن ترك حقل الاسم فارغًا")
                return
            try:
                with self.conn:
                    self.conn.execute("""
                        UPDATE trainees
                        SET name=?, rank=?, course=?, phone=?
                        WHERE national_id=?
                    """, (new_name, new_rank, new_course, new_phone, student[0]))
                messagebox.showinfo("نجاح", "تم التعديل بنجاح")
                edit_window.destroy()
                self.update_students_tree()
            except Exception as e:
                messagebox.showerror("خطأ", str(e))

        save_btn = tk.Button(button_frame, text="حفظ", font=self.fonts["text_bold"], bg=self.colors["success"],
                             fg="white",
                             padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=save_edit)
        save_btn.pack(side=tk.LEFT, padx=10)

        cancel_btn = tk.Button(button_frame, text="إلغاء", font=self.fonts["text_bold"], bg=self.colors["danger"],
                               fg="white",
                               padx=15, pady=5, bd=0, relief=tk.FLAT, cursor="hand2", command=edit_window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=10)

    def delete_selected_student(self):
        if not self.current_user["permissions"]["can_delete_students"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية حذف الطلاب")
            return

        selected_item = self.students_tree.selection()
        if not selected_item:
            messagebox.showinfo("تنبيه", "الرجاء تحديد طالب من القائمة")
            return
        values = self.students_tree.item(selected_item, "values")
        nid = values[0]
        if not messagebox.askyesnocancel("تأكيد", f"هل تريد حذف الطالب صاحب الهوية {nid}؟"):
            return
        try:
            with self.conn:
                self.conn.execute("DELETE FROM trainees WHERE national_id=?", (nid,))
            messagebox.showinfo("نجاح", "تم الحذف بنجاح")
            self.update_students_tree()
            self.update_statistics()
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def export_course_data(self, course_name):
        """وظيفة تصدير بيانات الدورة مع معلومات تفصيلية عن الغياب والتأخير"""
        if not self.current_user["permissions"]["can_export_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
            return

        try:
            # الحصول على بيانات الطلاب في الدورة
            cursor = self.conn.cursor()

            # استعلام البيانات الأساسية للطلاب
            cursor.execute("""
                SELECT national_id, name, rank, phone, is_excluded, exclusion_reason, excluded_date
                FROM trainees
                WHERE course=?
            """, (course_name,))
            students_data = cursor.fetchall()

            if not students_data:
                messagebox.showinfo("ملاحظة", f"لا يوجد طلاب مسجلين في الدورة '{course_name}'")
                return

            # إنشاء نافذة حالة لإظهار تقدم التصدير
            progress_window = tk.Toplevel(self.root)
            progress_window.title("جاري تصدير البيانات")
            progress_window.geometry("400x150")
            progress_window.configure(bg=self.colors["light"])
            progress_window.transient(self.root)
            progress_window.resizable(False, False)
            progress_window.grab_set()

            x = (progress_window.winfo_screenwidth() - 400) // 2
            y = (progress_window.winfo_screenheight() - 150) // 2
            progress_window.geometry(f"400x150+{x}+{y}")

            tk.Label(
                progress_window,
                text=f"جاري تصدير بيانات الدورة: {course_name}",
                font=self.fonts["text_bold"],
                bg=self.colors["light"],
                pady=10
            ).pack()

            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(
                progress_window,
                variable=progress_var,
                maximum=100,
                length=350
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

            # إنشاء قاموس لتخزين بيانات كل طالب
            students_dict = {}

            total_students = len(students_data)

            for index, student in enumerate(students_data):
                national_id, name, rank, phone, is_excluded, exclusion_reason, excluded_date = student

                # تحديث شريط التقدم
                progress_var.set((index / total_students) * 50)  # نصف التقدم للاستعلامات
                status_label.config(text=f"تحليل بيانات الطالب {index + 1} من {total_students}: {name}")
                progress_window.update()

                # إحصاء حالات الغياب والتأخير لكل طالب
                cursor.execute("""
                    SELECT status, COUNT(*) 
                    FROM attendance 
                    WHERE national_id=? AND status IN ('غائب', 'متأخر', 'غائب بعذر')
                    GROUP BY status
                """, (national_id,))
                status_counts = dict(cursor.fetchall())

                # الحصول على تواريخ الغياب
                cursor.execute("""
                    SELECT date FROM attendance 
                    WHERE national_id=? AND status='غائب'
                    ORDER BY date
                """, (national_id,))
                absent_dates = [row[0] for row in cursor.fetchall()]
                absent_dates_str = " || ".join(absent_dates) if absent_dates else ""

                # الحصول على تواريخ التأخير
                cursor.execute("""
                    SELECT date FROM attendance 
                    WHERE national_id=? AND status='متأخر'
                    ORDER BY date
                """, (national_id,))
                late_dates = [row[0] for row in cursor.fetchall()]
                late_dates_str = " || ".join(late_dates) if late_dates else ""

                # الحصول على أسباب وتواريخ الغياب بعذر
                cursor.execute("""
                    SELECT date, excuse_reason 
                    FROM attendance 
                    WHERE national_id=? AND status='غائب بعذر'
                    ORDER BY date
                """, (national_id,))
                excused_data = cursor.fetchall()

                excused_dates = [row[0] for row in excused_data]
                excused_dates_str = " || ".join(excused_dates) if excused_dates else ""

                # تجميع أسباب الغياب بعذر
                excuses_text = ""
                for date, reason in excused_data:
                    if reason:
                        excuses_text += f"{date}: {reason} || "

                if excuses_text.endswith(" || "):
                    excuses_text = excuses_text[:-4]

                # تنسيق حالة الاستبعاد
                excluded_status = "مستبعد" if is_excluded else "موجود"
                exclusion_info = exclusion_reason if is_excluded else "لا يوجد"
                exclusion_date = excluded_date if is_excluded else ""

                # تخزين البيانات في القاموس
                students_dict[national_id] = {
                    "رقم الهوية": national_id,
                    "الاسم": name,
                    "الرتبة": rank,
                    "الدورة": course_name,
                    "رقم الجوال": phone,
                    "عدد أيام الغياب": status_counts.get("غائب", 0),
                    "تواريخ الغياب": absent_dates_str,
                    "عدد أيام التأخير": status_counts.get("متأخر", 0),
                    "تواريخ التأخير": late_dates_str,
                    "عدد أيام الغياب بعذر": status_counts.get("غائب بعذر", 0),
                    "أسباب الغياب بعذر": excuses_text,
                    "حالة الطالب": excluded_status,
                    "سبب الاستبعاد": exclusion_info,
                    "تاريخ الاستبعاد": exclusion_date
                }

            # تحويل القاموس إلى DataFrame
            progress_var.set(60)
            status_label.config(text="إنشاء ملف التصدير...")
            progress_window.update()

            df = pd.DataFrame(list(students_dict.values()))

            # حفظ الإكسل
            progress_var.set(70)
            status_label.config(text="فتح حوار حفظ الملف...")
            progress_window.update()

            export_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"دورة_{course_name}.xlsx"
            )

            if not export_file:
                progress_window.destroy()
                return

            # تصدير إلى ملف إكسل مع معالجة خاصة للخلايا
            progress_var.set(80)
            status_label.config(text="إنشاء ملف Excel...")
            progress_window.update()

            writer = pd.ExcelWriter(export_file, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='بيانات الدورة')

            # الحصول على workbook وورقة العمل
            workbook = writer.book
            worksheet = writer.sheets['بيانات الدورة']

            # تنسيق الخلايا
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4F81BD',
                'font_color': 'white',
                'border': 1,
                'align': 'center'
            })

            # تنسيق للمحتوى
            content_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'top',
                'align': 'right',
                'border': 1
            })

            # تطبيق التنسيق على الرؤوس
            progress_var.set(90)
            status_label.config(text="تنسيق الملف...")
            progress_window.update()

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # تعديل عرض الأعمدة وتطبيق التنسيق
            for i, col in enumerate(df.columns):
                # تعيين عرض مناسب لكل عمود
                column_width = max(
                    df[col].astype(str).map(len).max(),  # أطول نص في العمود
                    len(str(col)) + 2  # عرض عنوان العمود + هامش
                )
                # تحديد حد أقصى
                if column_width > 50:
                    column_width = 50
                worksheet.set_column(i, i, column_width, content_format)

            # إضافة تنسيق لجدول البيانات (تصميم الجدول)
            worksheet.add_table(0, 0, len(df), len(df.columns) - 1, {
                'style': 'Table Style Medium 2',
                'columns': [{'header': col} for col in df.columns]
            })

            # حفظ الملف
            progress_var.set(95)
            status_label.config(text="حفظ الملف...")
            progress_window.update()

            writer.close()

            progress_var.set(100)
            status_label.config(text="تم تصدير البيانات بنجاح!")
            progress_window.update()

            # إغلاق نافذة التقدم بعد ثانيتين
            progress_window.after(2000, progress_window.destroy)

            messagebox.showinfo("نجاح", f"تم تصدير بيانات الدورة '{course_name}' بنجاح إلى الملف:\n{export_file}")

        except Exception as e:
            try:
                progress_window.destroy()
            except:
                pass
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير بيانات الدورة: {str(e)}")

    def import_new_course(self):
        """
        دالة استيراد دورة جديدة من ملف Excel
        مع إضافة تاريخ بداية ونهاية الدورة
        """
        if not self.current_user["permissions"]["can_import_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية استيراد البيانات")
            return

        import_window = tk.Toplevel(self.root)
        import_window.title("استيراد دورة جديدة")
        import_window.geometry("500x650")  # زيادة ارتفاع النافذة لاستيعاب حقول التاريخ
        import_window.configure(bg=self.colors["light"])
        import_window.transient(self.root)
        import_window.grab_set()

        x = (import_window.winfo_screenwidth() - 500) // 2
        y = (import_window.winfo_screenheight() - 650) // 2
        import_window.geometry(f"500x650+{x}+{y}")

        tk.Label(
            import_window,
            text="استيراد دورة جديدة",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        input_frame = tk.Frame(import_window, bg=self.colors["light"], padx=20, pady=20)
        input_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            input_frame,
            text="اسم الدورة الجديدة:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(anchor=tk.W, pady=(0, 5))

        course_entry = tk.Entry(input_frame, font=self.fonts["text"], width=40)
        course_entry.pack(fill=tk.X, pady=(0, 20))

        # إضافة إطار لتاريخ بداية الدورة
        start_date_frame = tk.LabelFrame(input_frame, text="تاريخ بداية الدورة", font=self.fonts["text_bold"],
                                         bg=self.colors["light"], padx=10, pady=10)
        start_date_frame.pack(fill=tk.X, pady=10)

        # حقول تاريخ البداية
        start_date_fields = tk.Frame(start_date_frame, bg=self.colors["light"])
        start_date_fields.pack(fill=tk.X)

        # اليوم
        tk.Label(start_date_fields, text="اليوم:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                          padx=5)
        start_day_entry = tk.Entry(start_date_fields, font=self.fonts["text"], width=5)
        start_day_entry.pack(side=tk.RIGHT, padx=5)

        # الشهر
        tk.Label(start_date_fields, text="الشهر:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                          padx=5)
        start_month_entry = tk.Entry(start_date_fields, font=self.fonts["text"], width=5)
        start_month_entry.pack(side=tk.RIGHT, padx=5)

        # السنة
        tk.Label(start_date_fields, text="السنة:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                          padx=5)
        start_year_entry = tk.Entry(start_date_fields, font=self.fonts["text"], width=8)
        start_year_entry.pack(side=tk.RIGHT, padx=5)

        # إضافة إطار لتاريخ نهاية الدورة
        end_date_frame = tk.LabelFrame(input_frame, text="تاريخ نهاية الدورة", font=self.fonts["text_bold"],
                                       bg=self.colors["light"], padx=10, pady=10)
        end_date_frame.pack(fill=tk.X, pady=10)

        # حقول تاريخ النهاية
        end_date_fields = tk.Frame(end_date_frame, bg=self.colors["light"])
        end_date_fields.pack(fill=tk.X)

        # اليوم
        tk.Label(end_date_fields, text="اليوم:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                        padx=5)
        end_day_entry = tk.Entry(end_date_fields, font=self.fonts["text"], width=5)
        end_day_entry.pack(side=tk.RIGHT, padx=5)

        # الشهر
        tk.Label(end_date_fields, text="الشهر:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                        padx=5)
        end_month_entry = tk.Entry(end_date_fields, font=self.fonts["text"], width=5)
        end_month_entry.pack(side=tk.RIGHT, padx=5)

        # السنة
        tk.Label(end_date_fields, text="السنة:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                        padx=5)
        end_year_entry = tk.Entry(end_date_fields, font=self.fonts["text"], width=8)
        end_year_entry.pack(side=tk.RIGHT, padx=5)

        # القيود على حقول التاريخ
        def validate_number(P, max_length):
            if P == "":
                return True
            if not P.isdigit():
                return False
            if len(P) > max_length:
                return False
            return True

        # تسجيل وظائف التحقق
        validate_day = import_window.register(lambda P: validate_number(P, 2))
        validate_month = import_window.register(lambda P: validate_number(P, 2))
        validate_year = import_window.register(lambda P: validate_number(P, 4))

        # تطبيق القيود
        start_day_entry.config(validate="key", validatecommand=(validate_day, "%P"))
        start_month_entry.config(validate="key", validatecommand=(validate_month, "%P"))
        start_year_entry.config(validate="key", validatecommand=(validate_year, "%P"))

        end_day_entry.config(validate="key", validatecommand=(validate_day, "%P"))
        end_month_entry.config(validate="key", validatecommand=(validate_month, "%P"))
        end_year_entry.config(validate="key", validatecommand=(validate_year, "%P"))

        columns_frame = tk.Frame(input_frame, bg=self.colors["light"])
        columns_frame.pack(fill=tk.X, pady=5)

        tk.Label(
            input_frame,
            text="الترتيب المفضل: (الاسم - الرتبة - رقم الهوية - رقم الجوال)",
            font=self.fonts["text"],
            bg=self.colors["light"],
            fg=self.colors["secondary"]
        ).pack(anchor=tk.W, pady=(0, 10))

        file_frame = tk.Frame(input_frame, bg=self.colors["light"])
        file_frame.pack(fill=tk.X)

        file_path_var = tk.StringVar()
        file_entry = tk.Entry(file_frame, textvariable=file_path_var, font=self.fonts["text"], width=30,
                              state="readonly")
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        def browse_file():
            file_path = filedialog.askopenfilename(
                title="اختر ملف Excel",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if file_path:
                file_path_var.set(file_path)

        browse_btn = tk.Button(
            file_frame,
            text="استعراض...",
            font=self.fonts["text"],
            bg=self.colors["secondary"],
            fg="white",
            padx=10, pady=3,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=browse_file
        )
        browse_btn.pack(side=tk.RIGHT)

        def import_course():
            course_name = course_entry.get().strip()
            file_path = file_path_var.get().strip()
            start_day = start_day_entry.get().strip()
            start_month = start_month_entry.get().strip()
            start_year = start_year_entry.get().strip()
            end_day = end_day_entry.get().strip()
            end_month = end_month_entry.get().strip()
            end_year = end_year_entry.get().strip()

            # التحقق من البيانات
            if not course_name:
                messagebox.showwarning("تنبيه", "الرجاء إدخال اسم الدورة")
                return

            if not file_path:
                messagebox.showwarning("تنبيه", "الرجاء اختيار ملف Excel")
                return

            # التحقق من تواريخ البداية والنهاية (اختياري)
            date_valid = True
            date_message = ""

            if (start_day or start_month or start_year) and not (start_day and start_month and start_year):
                date_valid = False
                date_message = "يجب إدخال تاريخ بداية الدورة كاملاً (اليوم والشهر والسنة)"

            if (end_day or end_month or end_year) and not (end_day and end_month and end_year):
                date_valid = False
                date_message = "يجب إدخال تاريخ نهاية الدورة كاملاً (اليوم والشهر والسنة)"

            if not date_valid:
                messagebox.showwarning("تنبيه", date_message)
                return

            # فحص الطلاب المتكررين واستكمال عملية الاستيراد
            # أضف استدعاء دالة جديدة تدعم حفظ التواريخ
            self.check_duplicate_students_with_dates(file_path, course_name,
                                                     start_day, start_month, start_year,
                                                     end_day, end_month, end_year)
            import_window.destroy()

        button_frame = tk.Frame(import_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=20)

        import_btn = tk.Button(
            button_frame,
            text="استيراد",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=import_course
        )
        import_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = tk.Button(
            button_frame,
            text="إلغاء",
            font=self.fonts["text_bold"],
            bg=self.colors["danger"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=import_window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)

    def check_duplicate_students_with_dates(self, file_path, course_name,
                                            start_day, start_month, start_year,
                                            end_day, end_month, end_year):
        """
        فحص الطلاب المتكررين قبل استيراد دورة جديدة مع دعم تواريخ البداية والنهاية
        """
        try:
            # إنشاء نافذة حالة لإظهار تقدم العملية
            progress_window = tk.Toplevel(self.root)
            progress_window.title("فحص الطلاب المتكررين")
            progress_window.geometry("400x150")
            progress_window.configure(bg=self.colors["light"])
            progress_window.transient(self.root)
            progress_window.grab_set()

            # توسيط النافذة
            x = (progress_window.winfo_screenwidth() - 400) // 2
            y = (progress_window.winfo_screenheight() - 150) // 2
            progress_window.geometry(f"400x150+{x}+{y}")

            tk.Label(
                progress_window,
                text="جاري فحص الطلاب المتكررين...",
                font=self.fonts["text_bold"],
                bg=self.colors["light"],
                pady=10
            ).pack()

            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(
                progress_window,
                variable=progress_var,
                maximum=100,
                length=350
            )
            progress_bar.pack(pady=10)

            status_label = tk.Label(
                progress_window,
                text="جاري قراءة ملف Excel...",
                font=self.fonts["text"],
                bg=self.colors["light"]
            )
            status_label.pack(pady=5)

            progress_window.update()

            # قراءة ملف Excel
            df = pd.read_excel(file_path)

            # تحديد الأعمدة المطلوبة (دعم الأسماء العربية والإنجليزية)
            column_mapping = {
                'الاسم': 'name',
                'رقم الهوية': 'national_id',
                'الرتبة': 'rank',
                'رقم الجوال': 'phone',
                'name': 'name',
                'national_id': 'national_id',
                'rank': 'rank',
                'phone': 'phone'
            }

            # تغيير أسماء الأعمدة إلى النموذج الإنجليزي
            df_columns = list(df.columns)
            english_columns = {}

            for col in df_columns:
                if col in column_mapping:
                    english_columns[col] = column_mapping[col]

            # التحقق من وجود الأعمدة المطلوبة
            required_cols_ar = ["الاسم", "رقم الهوية", "الرتبة", "رقم الجوال"]
            required_cols_en = ["name", "national_id", "rank", "phone"]

            # التحقق من وجود العمود بأي من اللغتين
            has_name = any(col in ["الاسم", "name"] for col in df_columns)
            has_id = any(col in ["رقم الهوية", "national_id"] for col in df_columns)
            has_rank = any(col in ["الرتبة", "rank"] for col in df_columns)
            has_phone = any(col in ["رقم الجوال", "phone"] for col in df_columns)

            if not (has_name and has_id):
                progress_window.destroy()
                messagebox.showwarning("تحذير",
                                       "يجب أن يحتوي الملف على الأعمدة التالية على الأقل:\n"
                                       "- الاسم (name)\n"
                                       "- رقم الهوية (national_id)")
                return False

            # إعادة تسمية الأعمدة للاستخدام الداخلي
            rename_dict = {}
            for orig_col in df.columns:
                if orig_col in column_mapping:
                    rename_dict[orig_col] = column_mapping[orig_col]

            if rename_dict:
                df = df.rename(columns=rename_dict)

            # إضافة الأعمدة المفقودة (اختياري) إذا لم تكن موجودة
            if 'rank' not in df.columns:
                df['rank'] = ''
            if 'phone' not in df.columns:
                df['phone'] = ''

            # قائمة الطلاب المتكررين
            duplicates = []

            # فحص كل طالب
            progress_var.set(20)
            status_label.config(text="جاري فحص الطلاب المتكررين...")
            progress_window.update()

            total_rows = len(df)
            cursor = self.conn.cursor()

            for i, row in enumerate(df.iterrows()):
                # تحديث شريط التقدم
                progress = 20 + (i / total_rows * 60)
                progress_var.set(progress)

                _, row_data = row
                # تحويل رقم الهوية إلى نص
                nid = str(row_data["national_id"]).strip()
                name = str(row_data["name"]).strip()

                if i % 10 == 0:
                    status_label.config(text=f"فحص الطالب {i + 1} من {total_rows}: {name}")
                    progress_window.update()

                # التحقق من وجود الطالب
                cursor.execute("""
                    SELECT t.course, t.name
                    FROM trainees t
                    WHERE t.national_id=?
                """, (nid,))

                result = cursor.fetchone()
                if result:
                    current_course, current_name = result
                    duplicates.append({
                        "id": nid,
                        "name": name,
                        "current_course": current_course
                    })

            progress_window.destroy()

            # حفظ تواريخ الدورة بغض النظر عن وجود طلاب متكررين
            self.save_course_dates(course_name, start_day, start_month, start_year, end_day, end_month, end_year)

            # عرض النتائج
            if duplicates:
                # عرض رسالة بأسماء الطلاب المتكررين فقط
                duplicate_details = f"تم العثور على {len(duplicates)} طالب موجودين بالفعل في دورات أخرى:\n\n"

                # عرض أول 10 طلاب فقط لتجنب رسائل طويلة جداً
                display_count = min(10, len(duplicates))
                for i in range(display_count):
                    duplicate_details += f"{i + 1}. {duplicates[i]['name']} (هوية: {duplicates[i]['id']}) - دورة: {duplicates[i]['current_course']}\n"

                if len(duplicates) > 10:
                    duplicate_details += f"\n... وغيرهم ({len(duplicates) - 10} آخرين)"

                duplicate_details += "\n\nهل تريد نقل هؤلاء الطلاب من دوراتهم السابقة إلى الدورة الجديدة؟"

                choice = messagebox.askquestion("طلاب متكررين", duplicate_details, type=messagebox.YESNOCANCEL)

                if choice == "cancel":
                    return False

                # متابعة الاستيراد مع خيار النقل (True) أو التخطي (False)
                update_mode = (choice == "yes")

                # إضافة سؤال عما إذا كانت الدورة متعددة الفصول
                is_multi_section = messagebox.askyesno("نوع الدورة", f"هل الدورة '{course_name}' متعددة الفصول؟")

                sections_count = 1
                if is_multi_section:
                    # طلب عدد الفصول
                    sections_count_str = simpledialog.askstring("عدد الفصول", "كم عدد الفصول في هذه الدورة؟",
                                                                initialvalue="2")
                    if not sections_count_str:
                        return False

                    try:
                        sections_count = int(sections_count_str)
                        if sections_count <= 0:
                            messagebox.showwarning("تنبيه", "يجب أن يكون عدد الفصول أكبر من صفر")
                            return False
                    except:
                        messagebox.showwarning("تنبيه", "الرجاء إدخال رقم صحيح لعدد الفصول")
                        return False
                else:
                    # إذا كانت الدورة غير متعددة الفصول، نجعلها بفصل واحد فقط
                    sections_count = 1
                    messagebox.showinfo("معلومات",
                                        f"سيتم إنشاء فصل واحد للدورة '{course_name}' ويمكنك إدارة الفصول لاحقًا من 'إدارة الفصول وتصدير الكشوفات'")

                # استدعاء دالة معالجة الاستيراد
                self.process_course_import_arabic(file_path, course_name, is_multi_section, sections_count, update_mode)
                return True
            else:
                messagebox.showinfo("تقرير الفحص",
                                    f"لم يتم العثور على طلاب متكررين. يمكنك المتابعة في استيراد الدورة '{course_name}'.")
                # إضافة سؤال عما إذا كانت الدورة متعددة الفصول
                is_multi_section = messagebox.askyesno("نوع الدورة", f"هل الدورة '{course_name}' متعددة الفصول؟")

                sections_count = 1
                if is_multi_section:
                    # طلب عدد الفصول
                    sections_count_str = simpledialog.askstring("عدد الفصول", "كم عدد الفصول في هذه الدورة؟",
                                                                initialvalue="2")
                    if not sections_count_str:
                        return False

                    try:
                        sections_count = int(sections_count_str)
                        if sections_count <= 0:
                            messagebox.showwarning("تنبيه", "يجب أن يكون عدد الفصول أكبر من صفر")
                            return False
                    except:
                        messagebox.showwarning("تنبيه", "الرجاء إدخال رقم صحيح لعدد الفصول")
                        return False
                else:
                    # إذا كانت الدورة غير متعددة الفصول، نجعلها بفصل واحد فقط
                    sections_count = 1
                    messagebox.showinfo("معلومات",
                                        f"سيتم إنشاء فصل واحد للدورة '{course_name}' ويمكنك إدارة الفصول لاحقًا من 'إدارة الفصول وتصدير الكشوفات'")

                # استدعاء دالة معالجة الاستيراد
                self.process_course_import_arabic(file_path, course_name, is_multi_section, sections_count, False)
                return False

        except Exception as e:
            try:
                progress_window.destroy()
            except:
                pass
            messagebox.showerror("خطأ", f"حدث خطأ أثناء فحص الطلاب المتكررين: {str(e)}")
            return False

    def save_course_dates(self, course_name, start_day, start_month, start_year, end_day, end_month, end_year):
        """
        حفظ تواريخ بداية ونهاية الدورة في قاعدة البيانات
        """
        try:
            current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with self.conn:
                # التحقق من وجود الدورة
                cursor = self.conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM course_info WHERE course_name=?", (course_name,))
                exists = cursor.fetchone()[0] > 0

                if exists:
                    # تحديث بيانات الدورة الموجودة
                    self.conn.execute("""
                        UPDATE course_info 
                        SET start_day=?, start_month=?, start_year=?, 
                            end_day=?, end_month=?, end_year=?
                        WHERE course_name=?
                    """, (start_day, start_month, start_year, end_day, end_month, end_year, course_name))
                else:
                    # إضافة دورة جديدة
                    self.conn.execute("""
                        INSERT INTO course_info 
                        (course_name, start_day, start_month, start_year, end_day, end_month, end_year, created_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (course_name, start_day, start_month, start_year, end_day, end_month, end_year, current_date))

            return True
        except Exception as e:
            print(f"خطأ في حفظ تواريخ الدورة: {str(e)}")
            return False

    def edit_course_dates(self, course_name):
        """تعديل تواريخ بداية ونهاية الدورة"""

        # استرجاع التواريخ الحالية من قاعدة البيانات
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT start_day, start_month, start_year, end_day, end_month, end_year
            FROM course_info
            WHERE course_name=?
        """, (course_name,))

        current_dates = cursor.fetchone() or ["", "", "", "", "", ""]

        # إنشاء نافذة تعديل التواريخ
        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"تعديل تواريخ دورة: {course_name}")
        edit_window.geometry("500x400")
        edit_window.configure(bg=self.colors["light"])
        edit_window.transient(self.root)
        edit_window.grab_set()

        # إضافة إطار لتاريخ بداية الدورة
        start_date_frame = tk.LabelFrame(edit_window, text="تاريخ بداية الدورة", font=self.fonts["text_bold"],
                                         bg=self.colors["light"], padx=10, pady=10)
        start_date_frame.pack(fill=tk.X, pady=10)

        # حقول تاريخ البداية
        start_date_fields = tk.Frame(start_date_frame, bg=self.colors["light"])
        start_date_fields.pack(fill=tk.X)

        # اليوم
        tk.Label(start_date_fields, text="اليوم:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                          padx=5)
        start_day_entry = tk.Entry(start_date_fields, font=self.fonts["text"], width=5)
        start_day_entry.insert(0, current_dates[0] if current_dates[0] else "")
        start_day_entry.pack(side=tk.RIGHT, padx=5)

        # الشهر
        tk.Label(start_date_fields, text="الشهر:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                          padx=5)
        start_month_entry = tk.Entry(start_date_fields, font=self.fonts["text"], width=5)
        start_month_entry.insert(0, current_dates[1] if current_dates[1] else "")
        start_month_entry.pack(side=tk.RIGHT, padx=5)

        # السنة
        tk.Label(start_date_fields, text="السنة:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                          padx=5)
        start_year_entry = tk.Entry(start_date_fields, font=self.fonts["text"], width=8)
        start_year_entry.insert(0, current_dates[2] if current_dates[2] else "")
        start_year_entry.pack(side=tk.RIGHT, padx=5)

        # نفس الشيء لتاريخ النهاية...
        end_date_frame = tk.LabelFrame(edit_window, text="تاريخ نهاية الدورة", font=self.fonts["text_bold"],
                                       bg=self.colors["light"], padx=10, pady=10)
        end_date_frame.pack(fill=tk.X, pady=10)

        # حقول تاريخ النهاية
        end_date_fields = tk.Frame(end_date_frame, bg=self.colors["light"])
        end_date_fields.pack(fill=tk.X)

        # اليوم
        tk.Label(end_date_fields, text="اليوم:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                        padx=5)
        end_day_entry = tk.Entry(end_date_fields, font=self.fonts["text"], width=5)
        end_day_entry.insert(0, current_dates[3] if current_dates[3] else "")
        end_day_entry.pack(side=tk.RIGHT, padx=5)

        # الشهر
        tk.Label(end_date_fields, text="الشهر:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                        padx=5)
        end_month_entry = tk.Entry(end_date_fields, font=self.fonts["text"], width=5)
        end_month_entry.insert(0, current_dates[4] if current_dates[4] else "")
        end_month_entry.pack(side=tk.RIGHT, padx=5)

        # السنة
        tk.Label(end_date_fields, text="السنة:", font=self.fonts["text"], bg=self.colors["light"]).pack(side=tk.RIGHT,
                                                                                                        padx=5)
        end_year_entry = tk.Entry(end_date_fields, font=self.fonts["text"], width=8)
        end_year_entry.insert(0, current_dates[5] if current_dates[5] else "")
        end_year_entry.pack(side=tk.RIGHT, padx=5)

        # أزرار الحفظ والإلغاء
        buttons_frame = tk.Frame(edit_window, bg=self.colors["light"], pady=20)
        buttons_frame.pack(fill=tk.X, padx=10)

        def save_dates():
            """حفظ التواريخ المعدلة"""
            start_day = start_day_entry.get().strip()
            start_month = start_month_entry.get().strip()
            start_year = start_year_entry.get().strip()
            end_day = end_day_entry.get().strip()
            end_month = end_month_entry.get().strip()
            end_year = end_year_entry.get().strip()

            # التحقق من صحة البيانات
            if (start_day or start_month or start_year) and not (start_day and start_month and start_year):
                messagebox.showwarning("تنبيه", "يجب إدخال تاريخ بداية الدورة كاملاً (اليوم والشهر والسنة)")
                return

            if (end_day or end_month or end_year) and not (end_day and end_month and end_year):
                messagebox.showwarning("تنبيه", "يجب إدخال تاريخ نهاية الدورة كاملاً (اليوم والشهر والسنة)")
                return

            # حفظ التواريخ الجديدة
            if self.save_course_dates(course_name, start_day, start_month, start_year, end_day, end_month, end_year):
                messagebox.showinfo("نجاح", "تم حفظ تواريخ الدورة بنجاح")
                edit_window.destroy()

        save_btn = tk.Button(
            buttons_frame,
            text="حفظ التغييرات",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=save_dates
        )
        save_btn.pack(side=tk.LEFT, padx=10)

        cancel_btn = tk.Button(
            buttons_frame,
            text="إلغاء",
            font=self.fonts["text_bold"],
            bg=self.colors["danger"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=edit_window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=10)

    def process_course_import_arabic(self, file_path, course_name, is_multi_section, sections_count, update_mode):
        """
        معالجة استيراد دورة من ملف Excel
        مع دعم الأعمدة باللغة العربية

        المعلمات:
            file_path (str): مسار ملف Excel
            course_name (str): اسم الدورة
            is_multi_section (bool): هل الدورة متعددة الفصول
            sections_count (int): عدد الفصول إذا كانت الدورة متعددة الفصول
            update_mode (bool): هل يتم نقل الطلاب من الدورات الأخرى (True) أو تخطيهم (False)
        """
        try:
            # إنشاء نافذة حالة لإظهار تقدم العملية
            progress_window = tk.Toplevel(self.root)
            progress_window.title("استيراد الدورة")
            progress_window.geometry("400x150")
            progress_window.configure(bg=self.colors["light"])
            progress_window.transient(self.root)
            progress_window.grab_set()

            # توسيط النافذة
            x = (progress_window.winfo_screenwidth() - 400) // 2
            y = (progress_window.winfo_screenheight() - 150) // 2
            progress_window.geometry(f"400x150+{x}+{y}")

            tk.Label(
                progress_window,
                text=f"جاري استيراد دورة '{course_name}'...",
                font=self.fonts["text_bold"],
                bg=self.colors["light"],
                pady=10
            ).pack()

            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(
                progress_window,
                variable=progress_var,
                maximum=100,
                length=350
            )
            progress_bar.pack(pady=10)

            status_label = tk.Label(
                progress_window,
                text="جاري قراءة ملف Excel...",
                font=self.fonts["text"],
                bg=self.colors["light"]
            )
            status_label.pack(pady=5)

            progress_window.update()

            # قراءة ملف Excel
            df = pd.read_excel(file_path)

            # الخطوة 1: التحقق من الأعمدة المطلوبة وتوحيد أسماء الأعمدة
            progress_var.set(10)
            status_label.config(text="التحقق من بنية البيانات...")
            progress_window.update()

            # تعريف جدول ترجمة الأعمدة من العربية إلى الإنجليزية
            column_mapping = {
                'الاسم': 'name',
                'رقم الهوية': 'national_id',
                'الرتبة': 'rank',
                'رقم الجوال': 'phone',
                'name': 'name',
                'national_id': 'national_id',
                'rank': 'rank',
                'phone': 'phone'
            }

            # التحقق مما إذا كانت الأعمدة المطلوبة موجودة (بالعربي أو الإنجليزي)
            has_name = any(col in ["الاسم", "name"] for col in df.columns)
            has_id = any(col in ["رقم الهوية", "national_id"] for col in df.columns)

            if not (has_name and has_id):
                progress_window.destroy()
                messagebox.showwarning("تحذير",
                                       "يجب أن يحتوي الملف على الأعمدة التالية على الأقل:\n"
                                       "- الاسم (name)\n"
                                       "- رقم الهوية (national_id)")
                return

            # إعادة تسمية الأعمدة للاستخدام الداخلي
            rename_dict = {}
            for orig_col in df.columns:
                if orig_col in column_mapping:
                    rename_dict[orig_col] = column_mapping[orig_col]

            if rename_dict:
                df = df.rename(columns=rename_dict)

            # إضافة الأعمدة المفقودة (اختياري) إذا لم تكن موجودة
            if 'rank' not in df.columns:
                df['rank'] = ''
            if 'phone' not in df.columns:
                df['phone'] = ''

            # إضافة عمود الدورة للطلاب
            df["course"] = course_name

            # الخطوة 2: التعامل مع الطلاب الموجودين
            progress_var.set(20)
            status_label.config(text="تحليل البيانات...")
            progress_window.update()

            cursor = self.conn.cursor()

            # متابعة اعداد الاحصائيات
            imported_count = 0
            skipped_count = 0
            moved_count = 0

            # قائمة لتخزين الطلاب المتخطين
            skipped_students = []

            # قائمة لتخزين بيانات الطلاب المستوردين
            imported_students = []

            # الخطوة 3: استيراد البيانات
            progress_var.set(30)
            status_label.config(text="جاري استيراد البيانات...")
            progress_window.update()

            total_rows = len(df)

            try:
                with self.conn:
                    for i, (_, row) in enumerate(df.iterrows()):
                        # تحديث شريط التقدم
                        progress = 30 + (i / total_rows * 40)  # من 30% إلى 70%
                        progress_var.set(progress)

                        if i % 10 == 0 or i == total_rows - 1:  # تحديث حالة التقدم
                            status_label.config(text=f"استيراد البيانات... ({i + 1}/{total_rows})")
                            progress_window.update()

                        # تأكد من تحويل البيانات إلى نصوص
                        nid = str(row["national_id"]).strip()
                        name = str(row["name"]).strip()
                        rank_ = str(row.get("rank", "")).strip()
                        phone = str(row.get("phone", "")).strip()

                        if not nid or not name:
                            skipped_count += 1
                            continue

                        # التحقق من وجود الطالب
                        cursor.execute("SELECT COUNT(*), course FROM trainees WHERE national_id=? GROUP BY course",
                                       (nid,))
                        result = cursor.fetchone()

                        if result:  # الطالب موجود بالفعل
                            exists, existing_course = result

                            if update_mode:  # وضع النقل
                                try:
                                    # حذف سجلات الحضور للطالب
                                    self.conn.execute("DELETE FROM attendance WHERE national_id=?", (nid,))
                                    # حذف الطالب من جدول الفصول إذا كان موجوداً
                                    self.conn.execute("DELETE FROM student_sections WHERE national_id=?", (nid,))
                                    # حذف الطالب
                                    self.conn.execute("DELETE FROM trainees WHERE national_id=?", (nid,))
                                    # إعادة إدخال الطالب بالدورة الجديدة
                                    self.conn.execute("""
                                        INSERT INTO trainees (national_id, name, rank, course, phone)
                                        VALUES (?, ?, ?, ?, ?)
                                    """, (nid, name, rank_, course_name, phone))
                                    moved_count += 1

                                    # إضافة الطالب إلى القائمة للتوزيع لاحقاً
                                    imported_students.append((nid, name))

                                except Exception as e:
                                    print(f"خطأ في نقل الطالب {nid}: {str(e)}")
                                    skipped_count += 1
                                    skipped_students.append({
                                        "id": nid,
                                        "name": name,
                                        "course": existing_course,
                                        "reason": f"خطأ أثناء النقل: {str(e)}"
                                    })
                            else:  # وضع التخطي
                                skipped_count += 1
                                skipped_students.append({
                                    "id": nid,
                                    "name": name,
                                    "course": existing_course,
                                    "reason": "موجود في دورة أخرى"
                                })
                            continue

                        # إضافة طالب جديد
                        self.conn.execute("""
                            INSERT INTO trainees (national_id, name, rank, course, phone)
                            VALUES (?, ?, ?, ?, ?)
                        """, (nid, name, rank_, course_name, phone))
                        imported_count += 1

                        # إضافة الطالب إلى القائمة للتوزيع لاحقاً
                        imported_students.append((nid, name))

                    # الخطوة 4: إنشاء الفصول وتوزيع الطلاب
                    if is_multi_section and (imported_count > 0 or moved_count > 0):
                        progress_var.set(75)
                        status_label.config(text="إنشاء الفصول وتوزيع الطلاب...")
                        progress_window.update()

                        current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        # إنشاء فصول بالعدد المطلوب
                        section_ids = []
                        for i in range(1, sections_count + 1):
                            section_name = f"فصل {i}"
                            self.conn.execute("""
                                INSERT INTO course_sections (course_name, section_name, created_date)
                                VALUES (?, ?, ?)
                            """, (course_name, section_name, current_date))

                            # الحصول على معرف الفصل المضاف
                            cursor.execute("SELECT last_insert_rowid()")
                            section_id = cursor.fetchone()[0]
                            section_ids.append((section_id, section_name))

                        # ترتيب قائمة الطلاب حسب الاسم
                        imported_students.sort(key=lambda x: x[1])

                        # حساب عدد الطلاب لكل فصل
                        total_students = len(imported_students)
                        students_per_section = total_students // sections_count
                        remainder = total_students % sections_count

                        # توزيع الطلاب على الفصول
                        current_index = 0
                        for i, (section_id, section_name) in enumerate(section_ids):
                            # حساب عدد الطلاب في هذا الفصل
                            section_student_count = students_per_section
                            if i < remainder:
                                section_student_count += 1

                            # إضافة الطلاب لهذا الفصل
                            for j in range(section_student_count):
                                if current_index < total_students:
                                    student_id, _ = imported_students[current_index]
                                    self.conn.execute("""
                                        INSERT OR REPLACE INTO student_sections
                                        (national_id, course_name, section_name, assigned_date)
                                        VALUES (?, ?, ?, ?)
                                    """, (student_id, course_name, section_name, current_date))
                                    current_index += 1

                    # التعديل الجديد: إنشاء فصل واحد افتراضي للدورات غير متعددة الفصول
                    elif (imported_count > 0 or moved_count > 0):  # إضافة هذا الشرط للدورات غير متعددة الفصول
                        # إنشاء فصل واحد افتراضي
                        progress_var.set(75)
                        status_label.config(text="إنشاء فصل واحد وتوزيع الطلاب...")
                        progress_window.update()

                        current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        section_name = "الفصل الرئيسي"  # أو "فصل 1" حسب تفضيلك

                        self.conn.execute("""
                            INSERT INTO course_sections (course_name, section_name, created_date)
                            VALUES (?, ?, ?)
                        """, (course_name, section_name, current_date))

                        # توزيع جميع الطلاب على الفصل الوحيد
                        for student_id, student_name in imported_students:
                            self.conn.execute("""
                                INSERT OR REPLACE INTO student_sections
                                (national_id, course_name, section_name, assigned_date)
                                VALUES (?, ?, ?, ?)
                            """, (student_id, course_name, section_name, current_date))

            except Exception as e:
                progress_window.destroy()
                messagebox.showerror("خطأ", f"حدث خطأ أثناء استيراد البيانات: {str(e)}")
                return

            # التحديث النهائي
            progress_var.set(95)
            status_label.config(text="إكمال عملية الاستيراد...")
            progress_window.update()

            # تحديث الإحصائيات
            self.update_statistics()
            self.update_students_tree()

            # إغلاق نافذة التقدم
            progress_window.destroy()

            # عرض نتائج الاستيراد
            result_message = f"تم استيراد دورة '{course_name}' بنجاح.\n\n"

            if imported_count > 0:
                result_message += f"• تم استيراد {imported_count} طالب جديد.\n"
            if moved_count > 0:
                result_message += f"• تم نقل {moved_count} طالب من دورات أخرى.\n"
            if skipped_count > 0:
                result_message += f"• تم تخطي {skipped_count} طالب.\n"

            if is_multi_section:
                result_message += f"\nتم إنشاء {sections_count} فصل للدورة وتوزيع الطلاب عليها بالترتيب الأبجدي."
            else:
                result_message += f"\nتم إنشاء فصل واحد للدورة (الفصل الرئيسي) وتوزيع الطلاب عليه."

            messagebox.showinfo("تقرير الاستيراد", result_message)

            # إذا كان هناك طلاب متخطون، عرض تفاصيلهم
            if skipped_students:
                skipped_details = "تفاصيل الطلاب المتخطين:\n\n"
                for i, student in enumerate(skipped_students, 1):
                    skipped_details += f"{i}. الاسم: {student['name']}, الهوية: {student['id']}\n"
                    skipped_details += f"   السبب: {student['reason']} - الدورة الحالية: {student['course']}\n"

                messagebox.showinfo("تفاصيل الطلاب المتخطين", skipped_details)

        except Exception as e:
            try:
                progress_window.destroy()
            except:
                pass
            messagebox.showerror("خطأ", f"حدث خطأ أثناء استيراد الدورة: {str(e)}")

    def view_student_profile(self):
        selected_item = self.students_tree.selection()
        if not selected_item:
            messagebox.showinfo("تنبيه", "الرجاء تحديد طالب من القائمة")
            return
        values = self.students_tree.item(selected_item, "values")
        if not values:
            return

    def view_student_profile(self):
        selected_item = self.students_tree.selection()
        if not selected_item:
            messagebox.showinfo("تنبيه", "الرجاء تحديد طالب من القائمة")
            return
        values = self.students_tree.item(selected_item, "values")
        if not values:
            return

        nid = values[0]
        name = values[1]

        # الحصول على معلومات الطالب بشكل منفصل
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM trainees WHERE national_id=?", (nid,))
        student = cursor.fetchone()

        if not student:
            messagebox.showinfo("تنبيه", "لا توجد معلومات عن هذا الطالب")
            return

        # الحصول على سجلات الحضور بشكل منفصل
        cursor.execute("""
            SELECT id, national_id, name, rank, course, time, date, status, original_status, 
                   registered_by, excuse_reason, updated_by, updated_at, modification_reason
            FROM attendance 
            WHERE national_id=?
            ORDER BY date DESC
        """, (nid,))
        attendance_records = cursor.fetchall()

        # تعريف المتغيرات الإحصائية مسبقًا بقيم افتراضية
        present_count = 0
        absent_count = 0
        late_count = 0
        excused_count = 0
        not_started_count = 0
        field_app_count = 0
        student_day_count = 0
        evening_remote_count = 0

        # حساب الإحصائيات مباشرة من قاعدة البيانات
        cursor.execute("""
            SELECT 
                COUNT(CASE WHEN status = 'حاضر' THEN 1 END) as present_count,
                COUNT(CASE WHEN status = 'غائب' THEN 1 END) as absent_count,
                COUNT(CASE WHEN status = 'متأخر' THEN 1 END) as late_count,
                COUNT(CASE WHEN status = 'غائب بعذر' THEN 1 END) as excused_count,
                COUNT(CASE WHEN status = 'لم يباشر' THEN 1 END) as not_started_count,
                COUNT(CASE WHEN status = 'تطبيق ميداني' THEN 1 END) as field_app_count,
                COUNT(CASE WHEN status = 'يوم طالب' THEN 1 END) as student_day_count,
                COUNT(CASE WHEN status = 'مسائية / عن بعد' THEN 1 END) as evening_remote_count,
                COUNT(CASE WHEN status = 'حالة وفاة' THEN 1 END) as death_case_count,
                COUNT(CASE WHEN status = 'منوم' THEN 1 END) as hospital_count
            FROM attendance
            WHERE national_id=?
        """, (nid,))

        stats = cursor.fetchone()
        if stats:
            present_count = stats[0] or 0
            absent_count = stats[1] or 0
            late_count = stats[2] or 0
            excused_count = stats[3] or 0
            not_started_count = stats[4] or 0
            field_app_count = stats[5] or 0
            student_day_count = stats[6] or 0
            evening_remote_count = stats[7] or 0
            death_case_count = stats[8] or 0  # إضافة جديدة
            hospital_count = stats[9] or 0  # إضافة جديدة

        # تصنيف السجلات حسب الحالة
        not_started_records = []
        absent_records = []
        late_records = []
        excused_records = []
        present_records = []
        field_application_records = []
        student_day_records = []
        evening_remote_records = []
        death_case_records = []  # قائمة لحالات الوفاة
        hospital_records = []  # قائمة للمنومين

        for record in attendance_records:
            status = record[7]  # حقل الحالة
            if status == "حاضر":
                present_records.append(record)
            elif status == "غائب":
                absent_records.append(record)
            elif status == "متأخر":
                late_records.append(record)
            elif status == "غائب بعذر":
                excused_records.append(record)
            elif status == "لم يباشر":
                not_started_records.append(record)
            elif status == "تطبيق ميداني":
                field_application_records.append(record)
            elif status == "يوم طالب":
                student_day_records.append(record)
            elif status == "مسائية / عن بعد":
                evening_remote_records.append(record)
            elif status == "حالة وفاة":
                death_case_records.append(record)
            elif status == "منوم":
                hospital_records.append(record)

        # إنشاء نافذة ملف الطالب
        profile_window = tk.Toplevel(self.root)
        profile_window.title(f"ملف الطالب - {name}")
        profile_window.geometry("1200x1200")  # زيادة الارتفاع
        profile_window.minsize(1000, 800)
        profile_window.configure(bg=self.colors["light"])
        profile_window.resizable(True, True)

        x = (profile_window.winfo_screenwidth() - 1200) // 2
        y = (profile_window.winfo_screenheight() - 1200) // 2
        profile_window.geometry(f"1200x1200+{x}+{y}")

        header_frame = tk.Frame(profile_window, bg=self.colors["primary"], padx=20, pady=15)
        header_frame.pack(fill=tk.X)

        # إضافة حالة الاستبعاد في العنوان إذا كان الطالب مستبعد
        if student[5] == 1:  # is_excluded
            exclusion_status = "- مستبعد"
            status_color = self.colors["excluded"]
        else:
            exclusion_status = ""
            status_color = self.colors["primary"]

        tk.Label(header_frame, text=f"ملف الطالب: {name} {exclusion_status}",
                 font=self.fonts["large_title"], bg=self.colors["primary"], fg="white").pack(anchor=tk.W)

        tk.Label(header_frame,
                 text=f"رقم الهوية: {nid} | الرتبة: {values[2]} | الدورة: {values[3]} | الجوال: {values[4]}",
                 font=self.fonts["text"], bg=self.colors["primary"], fg="white"
                 ).pack(anchor=tk.W, pady=(5, 0))

        # إضافة سبب الاستبعاد إذا كان الطالب مستبعد
        if student[5] == 1:  # is_excluded
            tk.Label(header_frame,
                     text=f"سبب الاستبعاد: {student[6]} | تاريخ الاستبعاد: {student[7]}",
                     font=self.fonts["text_bold"], bg=self.colors["primary"], fg="#ffcdd2"
                     ).pack(anchor=tk.W, pady=(5, 0))

        summary_frame = tk.Frame(profile_window, bg=self.colors["light"], padx=20, pady=15)
        summary_frame.pack(fill=tk.X)

        # إضافة زر استبعاد/إلغاء استبعاد الطالب في أعلى النافذة
        exclusion_frame = tk.Frame(summary_frame, bg=self.colors["light"], pady=5)
        exclusion_frame.pack(fill=tk.X)

        if student[5] == 1:  # الطالب مستبعد
            exclude_button = tk.Button(
                exclusion_frame,
                text="إلغاء استبعاد الطالب",
                font=self.fonts["text_bold"],
                bg=self.colors["success"],
                fg="white",
                padx=15, pady=5,
                bd=0, relief=tk.FLAT,
                cursor="hand2",
                command=lambda: self.toggle_student_exclusion(nid, False, profile_window)
            )
            exclude_button.pack(pady=10)
        else:
            exclude_button = tk.Button(
                exclusion_frame,
                text="استبعاد الطالب",
                font=self.fonts["text_bold"],
                bg=self.colors["excluded"],
                fg="white",
                padx=15, pady=5,
                bd=0, relief=tk.FLAT,
                cursor="hand2",
                command=lambda: self.toggle_student_exclusion(nid, True, profile_window)
            )
            exclude_button.pack(pady=10)

        stats_frame = tk.Frame(summary_frame, bg=self.colors["light"])
        stats_frame.pack(fill=tk.X)

        # استخدام الإحصائيات المحسوبة من قاعدة البيانات
        attendance_stats = [
            # السجلات الموجودة
            ("إجمالي أيام الحضور", present_count, self.colors["success"]),
            ("أيام التأخير", late_count, self.colors["late"]),
            ("أيام الغياب", absent_count, self.colors["danger"]),
            ("غياب بعذر", excused_count, self.colors["excused"]),
            ("لم يباشر", not_started_count, self.colors["not_started"]),
            ("تطبيق ميداني", field_app_count, self.colors["field_application"]),
            ("يوم طالب", student_day_count, self.colors["student_day"]),
            ("مسائية / عن بعد", evening_remote_count, self.colors["evening_remote"]),
            # إضافة الحالات الجديدة
            ("حالة وفاة", death_case_count, self.colors["death_case"]),
            ("منوم", hospital_count, self.colors["hospital"])
        ]

        for title, count, color in attendance_stats:
            stat_frame = tk.Frame(stats_frame, bg=self.colors["light"], bd=1, relief=tk.RIDGE, padx=5, pady=5)
            stat_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            tk.Label(stat_frame, text=title, font=self.fonts["text_bold"], bg=color, fg="white", padx=5, pady=5).pack(
                fill=tk.X)
            tk.Label(stat_frame, text=str(count), font=self.fonts["title"], bg=self.colors["light"]).pack(fill=tk.X,
                                                                                                          pady=5)

        details_notebook = ttk.Notebook(profile_window)
        details_notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        details_notebook = ttk.Notebook(profile_window)
        details_notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # إضافة تبويب المخالفات والإجراءات التأديبية
        violations_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(violations_frame, text="المخالفات والإجراءات التأديبية")

        # استرجاع سجل المخالفات الخاص بالطالب
        cursor.execute("""
            SELECT id, violation_date, violation_type, description, 
                   action_taken, action_date, recorded_by, notes
            FROM student_violations
            WHERE national_id=?
            ORDER BY violation_date DESC
        """, (nid,))
        violations_records = cursor.fetchall()

        if violations_records:

            # إضافة إطار للإحصائيات
            stats_frame = tk.Frame(violations_frame, bg=self.colors["light"])
            stats_frame.pack(fill=tk.X, pady=10)

            # حساب عدد المخالفات فقط
            violations_count = len(violations_records)

            # عرض إجمالي المخالفات فقط
            tk.Label(
                stats_frame,
                text="إحصائيات المخالفات:",
                font=self.fonts["text_bold"],
                bg=self.colors["light"],
                fg=self.colors["primary"]
            ).pack(anchor=tk.W, pady=(0, 5))

            cards_frame = tk.Frame(stats_frame, bg=self.colors["light"])
            cards_frame.pack(fill=tk.X)

            # بطاقة إجمالي المخالفات
            total_card = tk.Frame(cards_frame, bg=self.colors["light"], bd=1, relief=tk.RIDGE, padx=10, pady=5)
            total_card.pack(side=tk.RIGHT, padx=5, fill=tk.Y)

            tk.Label(
                total_card,
                text="إجمالي المخالفات",
                font=self.fonts["text_bold"],
                bg=self.colors["primary"],
                fg="white",
                padx=5, pady=5
            ).pack(fill=tk.X)

            tk.Label(
                total_card,
                text=str(violations_count),
                font=self.fonts["title"],
                bg=self.colors["light"]
            ).pack(fill=tk.X, pady=5)

            # إنشاء جدول لعرض المخالفات
            violations_tree = self.create_violations_table(violations_frame, violations_records)
        else:
            tk.Label(violations_frame, text="لا توجد مخالفات أو إجراءات تأديبية مسجلة",
                     font=self.fonts["subtitle"], bg=self.colors["light"], fg=self.colors["dark"],
                     pady=20).pack()

        # زر تسجيل مخالفة جديدة
        add_violation_btn = tk.Button(
            violations_frame,
            text="تسجيل مخالفة جديدة",
            font=self.fonts["text_bold"],
            bg=self.colors["warning"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.add_student_violation(nid, name, profile_window)
        )
        add_violation_btn.pack(pady=10)

        # دالات إنشاء الجداول اعتمادًا على صلاحيات المستخدم
        def create_attendance_detail_table(parent_frame, records):
            table_frame = tk.Frame(parent_frame, bg=self.colors["light"])
            table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            scrollbar = tk.Scrollbar(table_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # تحديد الأعمدة بناءً على صلاحيات المستخدم
            if self.current_user["permissions"]["can_view_edit_history"]:
                columns = ["date", "time", "registered_col", "updated_col", "mod_reason", "updated_when"]
            else:
                columns = ["date", "time", "registered_col"]

            detail_tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                       yscrollcommand=scrollbar.set, style="Profile.Treeview")

            detail_tree.heading("date", text="التاريخ")
            detail_tree.column("date", width=100, anchor=tk.CENTER)
            detail_tree.heading("time", text="الوقت")
            detail_tree.column("time", width=100, anchor=tk.CENTER)
            detail_tree.heading("registered_col", text="سجّل بواسطة")
            detail_tree.column("registered_col", width=180, anchor=tk.W)

            # إضافة أعمدة التعديل فقط للمستخدمين المصرح لهم
            if self.current_user["permissions"]["can_view_edit_history"]:
                detail_tree.heading("updated_col", text="من عدّل")
                detail_tree.column("updated_col", width=180, anchor=tk.W)
                detail_tree.heading("mod_reason", text="سبب التعديل")
                detail_tree.column("mod_reason", width=140, anchor=tk.W)
                detail_tree.heading("updated_when", text="وقت آخر تعديل")
                detail_tree.column("updated_when", width=150, anchor=tk.W)

            for rec in records:
                date_ = rec[6]
                time_ = rec[5]
                orig_status = rec[8] if rec[8] else rec[7]
                reg_by_ = rec[9]
                upd_by_ = rec[11] if rec[11] else ""
                upd_at_ = rec[12] if rec[12] else ""
                # تعديل مهم: عرض سبب التعديل فقط إذا كان هناك تعديل
                mod_reason_ = rec[13] if len(rec) > 13 and rec[13] and upd_by_ else ""
                registered_text = f"{reg_by_} (الحالة الاصلية: {orig_status})"

                if self.current_user["permissions"]["can_view_edit_history"]:
                    if upd_by_:
                        updated_text = f"{upd_by_} عدّلها إلى: {rec[7]}"
                    else:
                        updated_text = ""
                    detail_tree.insert("", tk.END,
                                       values=(date_, time_, registered_text, updated_text, mod_reason_, upd_at_))
                else:
                    detail_tree.insert("", tk.END, values=(date_, time_, registered_text))

            detail_tree.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=detail_tree.yview)

        def create_excused_detail_table(parent_frame, records):
            table_frame = tk.Frame(parent_frame, bg=self.colors["light"])
            table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            scrollbar = tk.Scrollbar(table_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # تحديد الأعمدة بناءً على صلاحيات المستخدم
            if self.current_user["permissions"]["can_view_edit_history"]:
                columns = ["date", "time", "registered_col", "excuse_reason", "updated_col", "mod_reason",
                           "updated_when"]
            else:
                columns = ["date", "time", "registered_col", "excuse_reason"]

            detail_tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                       yscrollcommand=scrollbar.set, style="Profile.Treeview")

            detail_tree.heading("date", text="التاريخ")
            detail_tree.column("date", width=100, anchor=tk.CENTER)
            detail_tree.heading("time", text="الوقت")
            detail_tree.column("time", width=100, anchor=tk.CENTER)
            detail_tree.heading("registered_col", text="سجّل بواسطة")
            detail_tree.column("registered_col", width=150, anchor=tk.W)
            detail_tree.heading("excuse_reason", text="سبب الغياب")
            detail_tree.column("excuse_reason", width=120, anchor=tk.W)

            # إضافة أعمدة التعديل فقط للمستخدمين المصرح لهم
            if self.current_user["permissions"]["can_view_edit_history"]:
                detail_tree.heading("updated_col", text="من عدّل")
                detail_tree.column("updated_col", width=120, anchor=tk.W)
                detail_tree.heading("mod_reason", text="سبب التعديل")
                detail_tree.column("mod_reason", width=140, anchor=tk.W)
                detail_tree.heading("updated_when", text="وقت آخر تعديل")
                detail_tree.column("updated_when", width=120, anchor=tk.W)

            for rec in records:
                date_ = rec[6]
                time_ = rec[5]
                orig_status = rec[8] if rec[8] else rec[7]
                reg_by_ = rec[9]
                excuse_ = rec[10] if rec[10] else "لا يوجد"
                upd_by_ = rec[11] if rec[11] else ""
                upd_at_ = rec[12] if rec[12] else ""
                # تعديل مهم: عرض سبب التعديل فقط إذا كان هناك تعديل (upd_by_ غير فارغ)
                mod_reason_ = rec[13] if len(rec) > 13 and rec[13] and upd_by_ else ""

                registered_text = f"{reg_by_} (الحالة الاصلية: {orig_status})"

                if self.current_user["permissions"]["can_view_edit_history"]:
                    if upd_by_:
                        updated_text = f"{upd_by_} عدّلها إلى: {rec[7]}"
                    else:
                        updated_text = ""
                    detail_tree.insert("", tk.END,
                                       values=(
                                           date_, time_, registered_text, excuse_, updated_text, mod_reason_, upd_at_))
                else:
                    detail_tree.insert("", tk.END, values=(date_, time_, registered_text, excuse_))

            detail_tree.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=detail_tree.yview)

        # تبويبات مفصلة
        not_started_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(not_started_frame, text=f"لم يباشر ({len(not_started_records)})")
        if not_started_records:
            create_attendance_detail_table(not_started_frame, not_started_records)
        else:
            tk.Label(not_started_frame, text="لا توجد أيام (لم يباشر)", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        absent_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(absent_frame, text=f"الغياب ({len(absent_records)})")
        if absent_records:
            create_attendance_detail_table(absent_frame, absent_records)
        else:
            tk.Label(absent_frame, text="لا توجد أيام غياب", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        late_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(late_frame, text=f"التأخير ({len(late_records)})")
        if late_records:
            create_attendance_detail_table(late_frame, late_records)
        else:
            tk.Label(late_frame, text="لا توجد أيام تأخير", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        excused_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(excused_frame, text=f"غياب بعذر ({len(excused_records)})")
        if excused_records:
            create_excused_detail_table(excused_frame, excused_records)
        else:
            tk.Label(excused_frame, text="لا توجد أيام غياب بعذر", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        present_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(present_frame, text=f"الحضور ({len(present_records)})")
        if present_records:
            create_attendance_detail_table(present_frame, present_records)
        else:
            tk.Label(present_frame, text="لا توجد أيام حضور", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        # إضافة تبويبات للحالات الجديدة
        field_app_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(field_app_frame, text=f"تطبيق ميداني ({len(field_application_records)})")
        if field_application_records:
            create_attendance_detail_table(field_app_frame, field_application_records)
        else:
            tk.Label(field_app_frame, text="لا توجد أيام تطبيق ميداني", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        student_day_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(student_day_frame, text=f"يوم طالب ({len(student_day_records)})")
        if student_day_records:
            create_attendance_detail_table(student_day_frame, student_day_records)
        else:
            tk.Label(student_day_frame, text="لا توجد أيام طالب", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        evening_remote_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(evening_remote_frame, text=f"مسائية / عن بعد ({len(evening_remote_records)})")
        if evening_remote_records:
            create_attendance_detail_table(evening_remote_frame, evening_remote_records)
        else:
            tk.Label(evening_remote_frame, text="لا توجد أيام مسائية / عن بعد", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()
        # بعد التبويبات الموجودة
        death_case_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(death_case_frame, text=f"حالة وفاة ({len(death_case_records)})")
        if death_case_records:
            create_excused_detail_table(death_case_frame, death_case_records)
        else:
            tk.Label(death_case_frame, text="لا توجد حالات وفاة", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        hospital_frame = tk.Frame(details_notebook, bg=self.colors["light"])
        details_notebook.add(hospital_frame, text=f"منوم ({len(hospital_records)})")
        if hospital_records:
            create_excused_detail_table(hospital_frame, hospital_records)
        else:
            tk.Label(hospital_frame, text="لا توجد حالات منوم", font=self.fonts["subtitle"],
                     bg=self.colors["light"], fg=self.colors["dark"], pady=20).pack()

        # إضافة إطار للأزرار
        button_frame = tk.Frame(profile_window, bg=self.colors["light"])
        button_frame.pack(pady=15)

        # زر تصدير بيانات الطالب
        export_btn = tk.Button(
            button_frame,
            text="تصدير بيانات الطالب",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=15, pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.export_student_to_word(student, attendance_records)
        )
        export_btn.pack(side=tk.LEFT, padx=10)

        # إضافة زر تصدير محاضر غيابات الطالب
        absence_report_btn = tk.Button(
            button_frame,
            text="تصدير محاضر غيابات الطالب",
            font=self.fonts["text_bold"],
            bg="#8E44AD",  # لون أرجواني
            fg="white",
            padx=15, pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.export_student_absence_reports(student, attendance_records)
        )
        absence_report_btn.pack(side=tk.LEFT, padx=10)

        # إضافة زر تصدير المخالفات
        violations_report_btn = tk.Button(
            button_frame,
            text="تصدير تقرير المخالفات",
            font=self.fonts["text_bold"],
            bg="#9C27B0",  # لون بنفسجي
            fg="white",
            padx=15, pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.export_student_violations_report(student, violations_records)
        )
        violations_report_btn.pack(side=tk.LEFT, padx=10)

        # زر الإغلاق
        close_btn = tk.Button(
            button_frame,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15, pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=profile_window.destroy
        )
        close_btn.pack(side=tk.LEFT, padx=10)

    def create_violations_table(self, parent_frame, violations_records):
        """إنشاء جدول لعرض مخالفات الطالب وإجراءاته التأديبية"""
        table_frame = tk.Frame(parent_frame, bg=self.colors["light"])
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # تحديد الأعمدة
        columns = ["date", "type", "desc", "action", "action_date", "recorded_by"]

        violations_tree = ttk.Treeview(table_frame, columns=columns,
                                       show="headings", yscrollcommand=scrollbar.set,
                                       style="Profile.Treeview")

        # تعريف العناوين
        violations_tree.heading("date", text="تاريخ المخالفة")
        violations_tree.column("date", width=100, anchor=tk.CENTER)

        violations_tree.heading("type", text="نوع المخالفة")
        violations_tree.column("type", width=120, anchor=tk.CENTER)

        violations_tree.heading("desc", text="وصف المخالفة")
        violations_tree.column("desc", width=180, anchor=tk.W)

        violations_tree.heading("action", text="الإجراء المتخذ")
        violations_tree.column("action", width=120, anchor=tk.CENTER)

        violations_tree.heading("action_date", text="تاريخ الإجراء")
        violations_tree.column("action_date", width=100, anchor=tk.CENTER)

        violations_tree.heading("recorded_by", text="المسجل")
        violations_tree.column("recorded_by", width=120, anchor=tk.W)

        # إضافة البيانات
        for rec in violations_records:
            violation_id = rec[0]
            violation_date = rec[1]
            violation_type = rec[2]
            description = rec[3]
            action_taken = rec[4]
            action_date = rec[5]
            recorded_by = rec[6]

            violations_tree.insert("", tk.END, values=(
                violation_date, violation_type, description,
                action_taken, action_date, recorded_by),
                                   tags=("violation",))

        # تنسيق صفوف الجدول
        violations_tree.tag_configure("violation", background="#FFF3E0")  # لون خلفية فاتح للمخالفات

        violations_tree.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=violations_tree.yview)

        # إضافة معالجة النقر المزدوج لعرض التفاصيل أو التعديل
        violations_tree.bind("<Double-1>", lambda event: self.view_violation_details(event, violations_tree))

        return violations_tree

    def add_student_violation(self, national_id, student_name, parent_window):
        """إضافة مخالفة جديدة للطالب"""
        if not self.current_user["permissions"]["can_edit_students"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل المخالفات")
            return

        violation_window = tk.Toplevel(parent_window)
        violation_window.title(f"تسجيل مخالفة جديدة - {student_name}")
        violation_window.geometry("700x600")
        violation_window.configure(bg=self.colors["light"])
        violation_window.transient(parent_window)
        violation_window.grab_set()

        # توسيط النافذة
        x = (violation_window.winfo_screenwidth() - 700) // 2
        y = (violation_window.winfo_screenheight() - 600) // 2
        violation_window.geometry(f"700x600+{x}+{y}")

        tk.Label(
            violation_window,
            text=f"تسجيل مخالفة جديدة للطالب: {student_name}",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        form_frame = tk.Frame(violation_window, bg=self.colors["light"], padx=20, pady=10)
        form_frame.pack(fill=tk.BOTH, expand=True)

        # نوع المخالفة
        tk.Label(form_frame, text="نوع المخالفة:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=0, column=1, sticky=tk.E, padx=5, pady=8)

        violation_types = [
            "تأخير متكرر",
            "غياب متكرر",
            "مخالفة تعليمات التدريب المستديمة",
            "عدم إنضباط",
            "أخرى"
        ]

        type_var = tk.StringVar()
        type_combo = ttk.Combobox(form_frame, textvariable=type_var, values=violation_types,
                                  font=self.fonts["text"], width=30)
        type_combo.grid(row=0, column=0, sticky=tk.W, padx=5, pady=8)

        # تاريخ المخالفة
        tk.Label(form_frame, text="تاريخ المخالفة:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=1, column=1, sticky=tk.E, padx=5, pady=8)

        violation_date = DateEntry(
            form_frame,
            width=15,
            background=self.colors["primary"],
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd',
            font=self.fonts["text"]
        )
        violation_date.grid(row=1, column=0, sticky=tk.W, padx=5, pady=8)

        # وصف المخالفة
        tk.Label(form_frame, text="وصف المخالفة:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=2, column=1, sticky=tk.NE, padx=5, pady=8)

        description_text = tk.Text(form_frame, font=self.fonts["text"], height=4, width=40)
        description_text.grid(row=2, column=0, sticky=tk.W, padx=5, pady=8)

        # الإجراء المتخذ
        tk.Label(form_frame, text="الإجراء المتخذ:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=3, column=1, sticky=tk.E, padx=5, pady=8)

        action_types = [
            "تعهد خطي",
            "إنذار",
            "توقيف",
            "أخرى"
        ]

        action_var = tk.StringVar()
        action_combo = ttk.Combobox(form_frame, textvariable=action_var, values=action_types,
                                    font=self.fonts["text"], width=30)
        action_combo.grid(row=3, column=0, sticky=tk.W, padx=5, pady=8)

        # تاريخ الإجراء
        tk.Label(form_frame, text="تاريخ الإجراء:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=4, column=1, sticky=tk.E, padx=5, pady=8)

        action_date = DateEntry(
            form_frame,
            width=15,
            background=self.colors["primary"],
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd',
            font=self.fonts["text"]
        )
        action_date.grid(row=4, column=0, sticky=tk.W, padx=5, pady=8)

        # ملاحظات إضافية
        tk.Label(form_frame, text="ملاحظات:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=5, column=1, sticky=tk.NE, padx=5, pady=8)

        notes_text = tk.Text(form_frame, font=self.fonts["text"], height=3, width=40)
        notes_text.grid(row=5, column=0, sticky=tk.W, padx=5, pady=8)

        # إضافة مرفق (اختياري)
        attachment_frame = tk.Frame(form_frame, bg=self.colors["light"])
        attachment_frame.grid(row=6, column=0, columnspan=2, sticky=tk.W, padx=5, pady=8)

        attachment_var = tk.StringVar()

        tk.Label(attachment_frame, text="مرفق (اختياري):",
                 font=self.fonts["text_bold"], bg=self.colors["light"]).pack(side=tk.RIGHT, padx=5)

        attachment_entry = tk.Entry(attachment_frame, textvariable=attachment_var,
                                    font=self.fonts["text"], width=30, state="readonly")
        attachment_entry.pack(side=tk.RIGHT, padx=5)

        def browse_file():
            file_path = filedialog.askopenfilename(
                title="اختر ملف المرفق",
                filetypes=[
                    ("PDF files", "*.pdf"),
                    ("Image files", "*.jpg *.jpeg *.png"),
                    ("All files", "*.*")
                ]
            )
            if file_path:
                attachment_var.set(file_path)

        browse_btn = tk.Button(
            attachment_frame,
            text="استعراض...",
            font=self.fonts["text"],
            bg=self.colors["secondary"],
            fg="white",
            padx=5, pady=2,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=browse_file
        )
        browse_btn.pack(side=tk.RIGHT, padx=5)

        # أزرار الحفظ والإلغاء
        button_frame = tk.Frame(violation_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=20)

        def save_violation():
            # التحقق من البيانات الإلزامية
            if not type_var.get() or not action_var.get():
                messagebox.showwarning("تنبيه", "يجب تحديد نوع المخالفة والإجراء المتخذ")
                return

            # استخلاص البيانات من النموذج
            v_type = type_var.get()
            v_date = violation_date.get_date().strftime("%Y-%m-%d")
            v_desc = description_text.get("1.0", tk.END).strip()
            a_type = action_var.get()
            a_date = action_date.get_date().strftime("%Y-%m-%d")
            notes = notes_text.get("1.0", tk.END).strip()
            attachment = attachment_var.get()

            # نسخ الملف المرفق إلى مجلد المرفقات إذا وجد
            attachment_path = ""
            if attachment:
                # إنشاء مجلد للمرفقات إذا لم يكن موجوداً
                attachments_dir = "attachments"
                if not os.path.exists(attachments_dir):
                    os.makedirs(attachments_dir)

                # نسخ الملف مع تسمية فريدة
                file_ext = os.path.splitext(attachment)[1]
                new_filename = f"{national_id}_{int(datetime.datetime.now().timestamp())}{file_ext}"
                new_path = os.path.join(attachments_dir, new_filename)

                try:
                    shutil.copy2(attachment, new_path)
                    attachment_path = new_path
                except Exception as e:
                    messagebox.showwarning("تنبيه", f"لم يتم نسخ الملف المرفق: {str(e)}")

            try:
                # حفظ المخالفة في قاعدة البيانات
                with self.conn:
                    self.conn.execute("""
                        INSERT INTO student_violations
                        (national_id, violation_date, violation_type, description, 
                         action_taken, action_date, recorded_by, notes, attachment_path)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        national_id, v_date, v_type, v_desc,
                        a_type, a_date, self.current_user["full_name"], notes, attachment_path
                    ))

                messagebox.showinfo("نجاح", "تم تسجيل المخالفة بنجاح")
                violation_window.destroy()

                # تحديث واجهة ملف الطالب (إعادة فتح الملف)
                parent_window.destroy()
                self.view_student_profile()

            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ المخالفة: {str(e)}")

        save_btn = tk.Button(
            button_frame,
            text="حفظ المخالفة",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=save_violation
        )
        save_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = tk.Button(
            button_frame,
            text="إلغاء",
            font=self.fonts["text_bold"],
            bg=self.colors["danger"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=violation_window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)

    def view_violation_details(self, event, tree):
        """عرض تفاصيل المخالفة وإمكانية تعديلها"""
        selected = tree.selection()
        if not selected:
            return

        item_values = tree.item(selected[0], "values")
        if not item_values:
            return

        # استرجاع معرف المخالفة
        violation_date = item_values[0]
        violation_type = item_values[1]

        # الحصول على معرف المخالفة من قاعدة البيانات
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT id, national_id, violation_date, violation_type, description, 
                   action_taken, action_date, recorded_by, notes, attachment_path
            FROM student_violations
            WHERE violation_date=? AND violation_type=?
            ORDER BY id DESC
            LIMIT 1
        """, (violation_date, violation_type))

        violation = cursor.fetchone()
        if not violation:
            return

        violation_id = violation[0]
        national_id = violation[1]

        # استرجاع اسم الطالب
        cursor.execute("SELECT name FROM trainees WHERE national_id=?", (national_id,))
        student_name = cursor.fetchone()[0]

        # عرض تفاصيل المخالفة
        details_window = tk.Toplevel(self.root)
        details_window.title(f"تفاصيل المخالفة - {student_name}")
        details_window.geometry("700x600")
        details_window.configure(bg=self.colors["light"])
        details_window.grab_set()

        # توسيط النافذة
        x = (details_window.winfo_screenwidth() - 700) // 2
        y = (details_window.winfo_screenheight() - 600) // 2
        details_window.geometry(f"700x600+{x}+{y}")

        tk.Label(
            details_window,
            text=f"تفاصيل المخالفة - {student_name}",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        details_frame = tk.Frame(details_window, bg=self.colors["light"], padx=20, pady=10)
        details_frame.pack(fill=tk.BOTH, expand=True)

        # عرض التفاصيل
        tk.Label(details_frame, text="نوع المخالفة:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=0, column=1, sticky=tk.E, padx=5, pady=8)
        tk.Label(details_frame, text=violation[3], font=self.fonts["text"],
                 bg=self.colors["light"]).grid(row=0, column=0, sticky=tk.W, padx=5, pady=8)

        tk.Label(details_frame, text="تاريخ المخالفة:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=1, column=1, sticky=tk.E, padx=5, pady=8)
        tk.Label(details_frame, text=violation[2], font=self.fonts["text"],
                 bg=self.colors["light"]).grid(row=1, column=0, sticky=tk.W, padx=5, pady=8)

        tk.Label(details_frame, text="وصف المخالفة:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=2, column=1, sticky=tk.NE, padx=5, pady=8)

        description_text = tk.Text(details_frame, font=self.fonts["text"], height=4, width=40)
        description_text.grid(row=2, column=0, sticky=tk.W, padx=5, pady=8)
        description_text.insert("1.0", violation[4])
        description_text.config(
            state=tk.DISABLED if not self.current_user["permissions"]["can_edit_students"] else tk.NORMAL)

        tk.Label(details_frame, text="الإجراء المتخذ:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=3, column=1, sticky=tk.E, padx=5, pady=8)
        tk.Label(details_frame, text=violation[5], font=self.fonts["text"],
                 bg=self.colors["light"]).grid(row=3, column=0, sticky=tk.W, padx=5, pady=8)

        tk.Label(details_frame, text="تاريخ الإجراء:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=4, column=1, sticky=tk.E, padx=5, pady=8)
        tk.Label(details_frame, text=violation[6], font=self.fonts["text"],
                 bg=self.colors["light"]).grid(row=4, column=0, sticky=tk.W, padx=5, pady=8)

        tk.Label(details_frame, text="المسجل:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=5, column=1, sticky=tk.E, padx=5, pady=8)
        tk.Label(details_frame, text=violation[7], font=self.fonts["text"],
                 bg=self.colors["light"]).grid(row=5, column=0, sticky=tk.W, padx=5, pady=8)

        tk.Label(details_frame, text="ملاحظات:", font=self.fonts["text_bold"],
                 bg=self.colors["light"]).grid(row=6, column=1, sticky=tk.NE, padx=5, pady=8)

        notes_text = tk.Text(details_frame, font=self.fonts["text"], height=3, width=40)
        notes_text.grid(row=6, column=0, sticky=tk.W, padx=5, pady=8)
        notes_text.insert("1.0", violation[8] if violation[8] else "")
        notes_text.config(state=tk.DISABLED if not self.current_user["permissions"]["can_edit_students"] else tk.NORMAL)

        # إذا كان هناك مرفق، إضافة زر لعرضه
        if violation[9]:
            attachment_frame = tk.Frame(details_frame, bg=self.colors["light"])
            attachment_frame.grid(row=7, column=0, columnspan=2, sticky=tk.W, padx=5, pady=8)

            tk.Label(attachment_frame, text="مرفق:", font=self.fonts["text_bold"],
                     bg=self.colors["light"]).pack(side=tk.RIGHT, padx=5)

            def open_attachment():
                try:
                    os.startfile(violation[9])
                except:
                    messagebox.showerror("خطأ", "لا يمكن فتح الملف المرفق")

            view_btn = tk.Button(
                attachment_frame,
                text="عرض المرفق",
                font=self.fonts["text"],
                bg=self.colors["secondary"],
                fg="white",
                padx=5, pady=2,
                bd=0, relief=tk.FLAT,
                cursor="hand2",
                command=open_attachment
            )
            view_btn.pack(side=tk.RIGHT, padx=5)

        # أزرار
        button_frame = tk.Frame(details_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=20)

        # إذا كان المستخدم لديه صلاحية التعديل، إضافة زر تحديث
        if self.current_user["permissions"]["can_edit_students"]:
            def update_violation():
                try:
                    new_desc = description_text.get("1.0", tk.END).strip()
                    new_notes = notes_text.get("1.0", tk.END).strip()

                    with self.conn:
                        self.conn.execute("""
                            UPDATE student_violations
                            SET description=?, notes=?
                            WHERE id=?
                        """, (new_desc, new_notes, violation_id))

                    messagebox.showinfo("نجاح", "تم تحديث تفاصيل المخالفة بنجاح")
                    details_window.destroy()

                    # تحديث واجهة ملف الطالب
                    self.view_student_profile()

                except Exception as e:
                    messagebox.showerror("خطأ", f"حدث خطأ أثناء تحديث المخالفة: {str(e)}")

            update_btn = tk.Button(
                button_frame,
                text="تحديث التفاصيل",
                font=self.fonts["text_bold"],
                bg=self.colors["success"],
                fg="white",
                padx=15, pady=5,
                bd=0, relief=tk.FLAT,
                cursor="hand2",
                command=update_violation
            )
            update_btn.pack(side=tk.LEFT, padx=5)

            # إضافة زر حذف المخالفة (للمشرفين فقط)
            if self.current_user["permissions"]["is_admin"]:
                def delete_violation():
                    if messagebox.askyesno("تأكيد الحذف",
                                           "هل أنت متأكد من حذف هذه المخالفة؟\nلا يمكن التراجع عن هذه العملية."):
                        try:
                            with self.conn:
                                self.conn.execute("DELETE FROM student_violations WHERE id=?", (violation_id,))

                            messagebox.showinfo("نجاح", "تم حذف المخالفة بنجاح")
                            details_window.destroy()

                            # تحديث واجهة ملف الطالب
                            self.view_student_profile()

                        except Exception as e:
                            messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف المخالفة: {str(e)}")

                delete_btn = tk.Button(
                    button_frame,
                    text="حذف المخالفة",
                    font=self.fonts["text_bold"],
                    bg=self.colors["danger"],
                    fg="white",
                    padx=15, pady=5,
                    bd=0, relief=tk.FLAT,
                    cursor="hand2",
                    command=delete_violation
                )
                delete_btn.pack(side=tk.LEFT, padx=5)

        close_btn = tk.Button(
            button_frame,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=details_window.destroy
        )
        close_btn.pack(side=tk.RIGHT, padx=5)

        # تعيين مؤشر التركيز على زر الإغلاق
        close_btn.focus_set()

    def export_student_violations_report(self, student_info, violations_records):
        """تصدير تقرير المخالفات والإجراءات التأديبية للطالب إلى ملف وورد"""
        if not self.current_user["permissions"]["can_export_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
            return

        try:
            # التأكد من وجود مكتبة python-docx
            if 'Document' not in globals():
                messagebox.showerror("خطأ",
                                     "لم يتم العثور على مكتبة python-docx. قم بتثبيتها باستخدام: pip install python-docx")
                return

            # استخراج بيانات الطالب
            nid = student_info[0]
            name = student_info[1]
            rank = student_info[2]
            course = student_info[3]

            # التحقق من وجود مخالفات
            if not violations_records or len(violations_records) == 0:
                messagebox.showinfo("معلومات", "لا توجد مخالفات مسجلة لهذا الطالب")
                return

            # إنشاء مستند جديد
            doc = Document()

            # إعداد المستند للغة العربية (RTL)
            section = doc.sections[0]
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)
            section.top_margin = Inches(1.0)
            section.bottom_margin = Inches(1.0)

            # إضافة عنوان المستند
            title = doc.add_heading('تقرير المخالفات والإجراءات التأديبية', level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in title.runs:
                run.font.size = Pt(18)
                run.font.bold = True
                run.font.rtl = True

            # إضافة بيانات الطالب
            info_para = doc.add_paragraph()
            info_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            info_para.add_run("بيانات الطالب:").bold = True

            info_table = doc.add_table(rows=1, cols=4)
            info_table.style = 'Table Grid'

            # إضافة رؤوس جدول بيانات الطالب
            header_cells = info_table.rows[0].cells

            # إضافة العناوين بشكل معكوس (RTL)
            header_cells[3].text = "الاسم"
            header_cells[2].text = "الرتبة"
            header_cells[1].text = "رقم الهوية"
            header_cells[0].text = "الدورة"

            # تنسيق الرؤوس
            for cell in header_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell.paragraphs[0].runs:
                    run.font.bold = True
                    run.font.rtl = True

            # إضافة بيانات الطالب
            data_cells = info_table.add_row().cells
            data_cells[3].text = name
            data_cells[2].text = rank
            data_cells[1].text = nid
            data_cells[0].text = course

            for cell in data_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell.paragraphs[0].runs:
                    run.font.rtl = True

            # إضافة فاصل قبل جدول المخالفات
            doc.add_paragraph()

            # عنوان جدول المخالفات
            violations_title = doc.add_paragraph()
            violations_title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            violations_title.add_run("سجل المخالفات:").bold = True

            # إنشاء جدول المخالفات
            violations_table = doc.add_table(rows=1, cols=5)
            violations_table.style = 'Table Grid'

            # رؤوس الجدول
            v_header_cells = violations_table.rows[0].cells
            v_header_cells[4].text = "تاريخ المخالفة"
            v_header_cells[3].text = "نوع المخالفة"
            v_header_cells[2].text = "الإجراء المتخذ"
            v_header_cells[1].text = "تاريخ الإجراء"
            v_header_cells[0].text = "وصف المخالفة"

            # تنسيق رؤوس الجدول
            for cell in v_header_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # إضافة تظليل للرأس
                try:
                    shading_elm = parse_xml(r'<w:shd {} w:fill="DDDDDD"/>'.format(nsdecls('w')))
                    cell._element.get_or_add_tcPr().append(shading_elm)
                except:
                    pass

                for run in cell.paragraphs[0].runs:
                    run.font.bold = True
                    run.font.rtl = True

            # إضافة بيانات المخالفات
            for violation in violations_records:
                # تاريخ المخالفة في [1]، نوع المخالفة في [2]، وصف المخالفة في [3]
                # الإجراء المتخذ في [4]، تاريخ الإجراء في [5]
                row_cells = violations_table.add_row().cells

                row_cells[4].text = violation[1]  # تاريخ المخالفة
                row_cells[3].text = violation[2]  # نوع المخالفة
                row_cells[2].text = violation[4]  # الإجراء المتخذ
                row_cells[1].text = violation[5]  # تاريخ الإجراء
                row_cells[0].text = violation[3]  # وصف المخالفة

                # تنسيق الخلايا
                for i, cell in enumerate(row_cells):
                    alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # جعل خلية الوصف محاذاة يمين
                    if i == 0:
                        alignment = WD_ALIGN_PARAGRAPH.RIGHT

                    cell.paragraphs[0].alignment = alignment
                    for run in cell.paragraphs[0].runs:
                        run.font.rtl = True

            # إضافة بيانات كل مخالفة بالتفصيل في صفحات منفصلة
            for i, violation in enumerate(violations_records):
                # إضافة صفحة جديدة
                doc.add_page_break()

                # عنوان المخالفة
                v_title = doc.add_heading(f'تفاصيل المخالفة رقم {i + 1}', level=1)
                v_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in v_title.runs:
                    run.font.rtl = True

                # تفاصيل المخالفة
                details_table = doc.add_table(rows=7, cols=2)
                details_table.style = 'Table Grid'

                # إضافة البيانات
                details_table.cell(0, 1).text = "نوع المخالفة"
                details_table.cell(0, 0).text = violation[2]

                details_table.cell(1, 1).text = "تاريخ المخالفة"
                details_table.cell(1, 0).text = violation[1]

                details_table.cell(2, 1).text = "وصف المخالفة"
                details_table.cell(2, 0).text = violation[3]

                details_table.cell(3, 1).text = "الإجراء المتخذ"
                details_table.cell(3, 0).text = violation[4]

                details_table.cell(4, 1).text = "تاريخ الإجراء"
                details_table.cell(4, 0).text = violation[5]

                details_table.cell(5, 1).text = "المسجل"
                details_table.cell(5, 0).text = violation[6]

                details_table.cell(6, 1).text = "ملاحظات"
                details_table.cell(6, 0).text = violation[7] if violation[7] else "لا توجد ملاحظات"

                # تنسيق الجدول
                for row in details_table.rows:
                    # خلية العنوان
                    row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in row.cells[1].paragraphs[0].runs:
                        run.font.bold = True
                        run.font.rtl = True

                    # تطبيق تظليل على خلايا العناوين
                    try:
                        shading_elm = parse_xml(r'<w:shd {} w:fill="E9E9E9"/>'.format(nsdecls('w')))
                        row.cells[1]._element.get_or_add_tcPr().append(shading_elm)
                    except:
                        pass

                    # خلية المحتوى
                    row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    for run in row.cells[0].paragraphs[0].runs:
                        run.font.rtl = True

            # إضافة تذييل بتاريخ الطباعة
            doc.add_paragraph()
            footer_para = doc.add_paragraph()
            footer_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            today_date = datetime.datetime.now().strftime("%Y-%m-%d")
            footer_text = footer_para.add_run(
                f"تاريخ الطباعة: {today_date} - طُبع بواسطة: {self.current_user['full_name']}")
            footer_text.font.size = Pt(9)
            footer_text.font.rtl = True

            # حفظ المستند
            export_file = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile=f"تقرير_مخالفات_{name}_{nid}.docx"
            )

            if export_file:
                doc.save(export_file)
                messagebox.showinfo("نجاح", f"تم تصدير تقرير المخالفات بنجاح إلى:\n{export_file}")

                # محاولة فتح الملف
                try:
                    os.startfile(export_file)
                except:
                    pass

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير تقرير المخالفات: {str(e)}")

    def toggle_student_exclusion(self, national_id, exclude, profile_window=None):
        """إضافة أو إزالة استبعاد الطالب"""
        if not self.current_user["permissions"]["can_edit_students"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية استبعاد الطلاب")
            return

        cursor = self.conn.cursor()
        cursor.execute("SELECT name, is_excluded FROM trainees WHERE national_id=?", (national_id,))
        student = cursor.fetchone()

        if not student:
            messagebox.showwarning("تنبيه", "لم يتم العثور على الطالب")
            return

        student_name, current_excluded = student

        if exclude and current_excluded == 1:
            messagebox.showinfo("تنبيه", "هذا الطالب مستبعد بالفعل")
            return

        if not exclude and current_excluded == 0:
            messagebox.showinfo("تنبيه", "هذا الطالب غير مستبعد بالفعل")
            return

        if exclude:
            reason = simpledialog.askstring("سبب الاستبعاد", "أدخل سبب استبعاد الطالب:")
            if reason is None:  # الضغط على إلغاء
                return

            current_date = datetime.datetime.now().strftime("%Y-%m-%d")

            try:
                with self.conn:
                    self.conn.execute("""
                        UPDATE trainees 
                        SET is_excluded=1, exclusion_reason=?, excluded_date=? 
                        WHERE national_id=?
                    """, (reason, current_date, national_id))

                messagebox.showinfo("نجاح", f"تم استبعاد الطالب {student_name} بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء استبعاد الطالب: {str(e)}")
                return
        else:
            try:
                with self.conn:
                    self.conn.execute("""
                        UPDATE trainees 
                        SET is_excluded=0, exclusion_reason='', excluded_date='' 
                        WHERE national_id=?
                    """, (national_id,))

                messagebox.showinfo("نجاح", f"تم إلغاء استبعاد الطالب {student_name} بنجاح")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء إلغاء استبعاد الطالب: {str(e)}")
                return

        # تحديث الإحصائيات والبيانات
        self.update_students_tree()
        self.update_statistics()
        self.update_attendance_display()

        # إغلاق نافذة ملف الطالب إذا كانت مفتوحة وإعادة فتحها لتعكس التغييرات
        if profile_window:
            profile_window.destroy()
            self.view_student_profile()

    def export_course_completion(self):
        """وظيفة تصدير مستند تكميل الدورات بتنسيق Word مع ملخص للدورات في الصفحة الأولى مع إضافة تواريخ بداية ونهاية الدورة"""
        if not self.current_user["permissions"]["can_export_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
            return

        try:
            # التأكد من وجود مكتبة python-docx
            if 'Document' not in globals():
                messagebox.showerror("خطأ",
                                     "لم يتم العثور على مكتبة python-docx. قم بتثبيتها باستخدام: pip install python-docx")
                return

            # الحصول على تاريخ اليوم المحدد
            selected_date = self.log_date_entry.get_date()
            selected_date_str = selected_date.strftime("%Y-%m-%d")
            # تنسيق التاريخ بشكل أفضل للعرض
            arabic_date = selected_date.strftime("%Y/%m/%d")

            # تحديد يوم الأسبوع بالعربية
            weekday = selected_date.weekday()
            arabic_weekdays = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
            arabic_weekday = arabic_weekdays[weekday]

            # استعلام إجمالي عدد الطلاب في النظام (غير المستبعدين)
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT COUNT(*)
                FROM trainees
                WHERE is_excluded=0
            """)
            total_students_count = cursor.fetchone()[0]

            # الحصول على عدد الطلاب غير المسجلين بعد
            cursor.execute("""
                SELECT COUNT(*)
                FROM trainees t
                WHERE t.is_excluded=0 AND NOT EXISTS (
                    SELECT 1 FROM attendance a
                    WHERE a.national_id = t.national_id AND a.date = ?
                )
            """, (selected_date_str,))

            unrecorded_count = cursor.fetchone()[0]

            # التحقق من عدم وجود طلاب غير مسجلين (وفقاً لطلب المستخدم)
            if unrecorded_count > 0:
                messagebox.showwarning("تنبيه",
                                       f"لا يمكن تصدير التكميل، هناك {unrecorded_count} طالب لم يتم تسجيل حضورهم/غيابهم بعد.")
                return

            # استعلام بيانات الطلاب الغائبين والذين لم يباشروا والغائبين بعذر وجميع الحالات الأخرى
            cursor.execute("""
                SELECT a.national_id, a.name, a.rank, a.course, a.status, a.excuse_reason
                FROM attendance a
                JOIN trainees t ON a.national_id = t.national_id
                WHERE a.date=? AND t.is_excluded=0 AND a.status IN ('غائب', 'لم يباشر', 'غائب بعذر', 'متأخر', 'تطبيق ميداني', 'يوم طالب', 'مسائية / عن بعد', 'حالة وفاة', 'منوم')
            """, (selected_date_str,))
            all_attendance_data = cursor.fetchall()

            if not all_attendance_data and total_students_count == 0:
                messagebox.showinfo("ملاحظة", "لا توجد بيانات غياب أو طلاب في النظام لهذا اليوم.")
                return

            # الحصول على إحصائيات الدورات ولكن فقط التي لها قيم (لا نريد أعمدة صفرية)
            cursor.execute("""
                SELECT 
                    t.course,
                    COUNT(DISTINCT t.national_id) as total_course_students,
                    COUNT(DISTINCT CASE WHEN a.status = 'حاضر' THEN a.national_id END) as present_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'غائب' THEN a.national_id END) as absent_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'غائب بعذر' THEN a.national_id END) as excused_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'متأخر' THEN a.national_id END) as late_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'لم يباشر' THEN a.national_id END) as not_started_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'تطبيق ميداني' THEN a.national_id END) as field_app_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'يوم طالب' THEN a.national_id END) as student_day_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'مسائية / عن بعد' THEN a.national_id END) as evening_remote_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'حالة وفاة' THEN a.national_id END) as death_case_count,
                    COUNT(DISTINCT CASE WHEN a.status = 'منوم' THEN a.national_id END) as hospital_count,
                    COUNT(DISTINCT CASE WHEN a.status IS NOT NULL THEN a.national_id END) as recorded_count
                FROM 
                    trainees t
                LEFT JOIN 
                    attendance a ON t.national_id = a.national_id AND a.date = ?
                WHERE 
                    t.is_excluded = 0
                GROUP BY 
                    t.course
            """, (selected_date_str,))

            courses_stats = cursor.fetchall()

            if not courses_stats:
                messagebox.showinfo("ملاحظة", "لا توجد دورات مسجلة في النظام.")
                return

            # تصنيف البيانات حسب الحالة
            absent_data = [student for student in all_attendance_data if student[4] == 'غائب']
            not_started_data = [student for student in all_attendance_data if student[4] == 'لم يباشر']
            excused_data = [student for student in all_attendance_data if student[4] == 'غائب بعذر']
            late_data = [student for student in all_attendance_data if student[4] == 'متأخر']
            field_app_data = [student for student in all_attendance_data if student[4] == 'تطبيق ميداني']
            student_day_data = [student for student in all_attendance_data if student[4] == 'يوم طالب']
            evening_remote_data = [student for student in all_attendance_data if student[4] == 'مسائية / عن بعد']
            death_case_data = [student for student in all_attendance_data if student[4] == 'حالة وفاة']
            hospital_data = [student for student in all_attendance_data if student[4] == 'منوم']

            # استعلام إحصائيات الحضور الإجمالية لهذا اليوم
            cursor.execute("""
                SELECT 
                    COUNT(CASE WHEN a.status = 'حاضر' THEN 1 END) as present_count,
                    COUNT(CASE WHEN a.status = 'غائب' THEN 1 END) as absent_count,
                    COUNT(CASE WHEN a.status = 'غائب بعذر' THEN 1 END) as excused_count,
                    COUNT(CASE WHEN a.status = 'متأخر' THEN 1 END) as late_count,
                    COUNT(CASE WHEN a.status = 'لم يباشر' THEN 1 END) as not_started_count,
                    COUNT(CASE WHEN a.status = 'تطبيق ميداني' THEN 1 END) as field_app_count,
                    COUNT(CASE WHEN a.status = 'يوم طالب' THEN 1 END) as student_day_count,
                    COUNT(CASE WHEN a.status = 'مسائية / عن بعد' THEN 1 END) as evening_remote_count,
                    COUNT(CASE WHEN a.status = 'حالة وفاة' THEN 1 END) as death_case_count,
                    COUNT(CASE WHEN a.status = 'منوم' THEN 1 END) as hospital_count,
                    COUNT(*) as total_recorded_count
                FROM attendance a
                JOIN trainees t ON a.national_id = t.national_id
                WHERE a.date=? AND t.is_excluded=0
            """, (selected_date_str,))
            stats = cursor.fetchone()

            # إذا لم تتوفر إحصائيات (لا يتوقع حدوث ذلك)
            if not stats:
                stats = (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)  # قيم افتراضية شاملة الحالات الجديدة

            # أعداد الحضور والغياب
            present_count = stats[0] or 0
            absent_count = stats[1] or 0
            excused_count = stats[2] or 0
            late_count = stats[3] or 0
            not_started_count = stats[4] or 0
            field_app_count = stats[5] or 0
            student_day_count = stats[6] or 0
            evening_remote_count = stats[7] or 0
            death_case_count = stats[8] or 0
            hospital_count = stats[9] or 0
            total_recorded_count = stats[10] or 0

            # تحديد العناوين التي سيتم عرضها في الجدول بناءً على وجود قيم
            # (نعرض العمود فقط إذا كانت هناك حالة واحدة على الأقل)
            show_field_app = any(course[7] > 0 for course in courses_stats)
            show_student_day = any(course[8] > 0 for course in courses_stats)
            show_evening_remote = any(course[9] > 0 for course in courses_stats)
            show_death_case = any(course[10] > 0 for course in courses_stats)
            show_hospital = any(course[11] > 0 for course in courses_stats)

            # إنشاء مستند Word جديد
            doc = Document()

            # إعداد صفحة المستند
            section = doc.sections[0]
            section.page_width = Inches(11.69)  # A4 width in landscape
            section.page_height = Inches(8.27)  # A4 height in landscape
            section.orientation = WD_ORIENTATION.LANDSCAPE
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            section.top_margin = Inches(0.7)
            section.bottom_margin = Inches(0.7)

            # إضافة العنوان المعدل
            title = doc.add_heading(
                f'التكميل اليومي لدورات التخصصية المنعقدة بمدينة تدريب الامن العام بالمنطقة الشرقية ليوم {arabic_weekday} بتاريخ {arabic_date}',
                level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in title.runs:
                run.font.rtl = True
                run.font.bold = True
                run.font.size = Pt(16)

            # إضافة خط أفقي فاصل
            border_para = doc.add_paragraph()
            border_para.paragraph_format.border_bottom = True

            # إضافة فقرة اسم الجهة
            dept_para = doc.add_paragraph()
            dept_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            dept_run = dept_para.add_run("قسم شؤون المدربين")
            dept_run.font.rtl = True
            dept_run.font.bold = True
            dept_run.font.size = Pt(14)

            # إضافة فقرة فاصلة
            doc.add_paragraph()

            # حساب عدد الأعمدة المطلوبة بناءً على الحالات الموجودة
            # 9 أعمدة ثابتة ثم نضيف عمود لكل حالة موجودة
            cols_count = 9  # العدد، الدورة، بداية، نهاية، القوة، لم يباشر، غياب، تأخير، غياب بعذر

            # إضافة أعمدة إضافية للحالات الموجودة فقط
            if show_field_app:
                cols_count += 1
            if show_student_day:
                cols_count += 1
            if show_evening_remote:
                cols_count += 1
            if show_death_case:
                cols_count += 1
            if show_hospital:
                cols_count += 1

            # إضافة عمود العدد الفعلي دائماً
            cols_count += 1

            # إنشاء جدول ملخص الدورات في الصفحة الأولى مع أعمدة إضافية لتاريخ بداية ونهاية الدورة
            courses_table = doc.add_table(rows=1, cols=cols_count)
            courses_table.style = 'Table Grid'

            # إنشاء قائمة العناوين
            headers = [
                "العدد", "اسم الدورة",
                "تاريخ بداية الدورة", "تاريخ نهاية الدورة",
                "القوة", "لم يباشر", "غياب", "تأخير", "غياب بعذر"
            ]

            # إضافة عناوين الحالات الموجودة فقط
            if show_field_app:
                headers.append("تطبيق ميداني")
            if show_student_day:
                headers.append("يوم طالب")
            if show_evening_remote:
                headers.append("مسائية / عن بعد")
            if show_death_case:
                headers.append("حالة وفاة")
            if show_hospital:
                headers.append("منوم")

            # إضافة عنوان العدد الفعلي دائماً
            headers.append("العدد الفعلي")

            # عناوين جدول الدورات
            header_cells = courses_table.rows[0].cells

            for i, header in enumerate(headers):
                # حساب الموقع المناسب للعناوين (من اليمين إلى اليسار بسبب RTL)
                idx = len(headers) - i - 1
                header_cells[idx].text = header
                header_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                for run in header_cells[idx].paragraphs[0].runs:
                    run.font.bold = True
                    run.font.rtl = True
                    run.font.size = Pt(11)

                # تطبيق تظليل للرأس
                try:
                    shading_elm = parse_xml(r'<w:shd {} w:fill="DDDDDD"/>'.format(nsdecls('w')))
                    header_cells[idx]._element.get_or_add_tcPr().append(shading_elm)
                except:
                    pass  # تجاهل الخطأ إذا حدث

            # إضافة بيانات الدورات إلى الجدول
            for idx, course_data in enumerate(courses_stats):
                course_name = course_data[0]
                total_course_students = course_data[1] or 0
                present_count = course_data[2] or 0
                absent_count = course_data[3] or 0
                excused_count = course_data[4] or 0
                late_count = course_data[5] or 0
                not_started_count = course_data[6] or 0
                field_app_count = course_data[7] or 0
                student_day_count = course_data[8] or 0
                evening_remote_count = course_data[9] or 0
                death_case_count = course_data[10] or 0
                hospital_count = course_data[11] or 0
                recorded_count = course_data[12] or 0

                # الحصول على تواريخ البداية والنهاية للدورة من قاعدة البيانات
                cursor.execute("""
                    SELECT start_day, start_month, start_year, end_day, end_month, end_year
                    FROM course_info
                    WHERE course_name=?
                """, (course_name,))
                date_info = cursor.fetchone()

                start_date_str = ""
                end_date_str = ""

                if date_info:
                    start_day, start_month, start_year, end_day, end_month, end_year = date_info
                    if start_day and start_month and start_year:
                        start_date_str = f"{start_day}/{start_month}/{start_year}"
                    if end_day and end_month and end_year:
                        end_date_str = f"{end_day}/{end_month}/{end_year}"

                # حساب العدد الفعلي (بعد خصم جميع الحالات المذكورة)
                # العدد الفعلي = عدد الحاضرين فقط
                effective_count = present_count

                # حساب المجموع
                total_status_count = (absent_count + excused_count + late_count + not_started_count +
                                      field_app_count + student_day_count + evening_remote_count +
                                      death_case_count + hospital_count + present_count)

                row_cells = courses_table.add_row().cells

                # تحضير القيم للإدخال في الجدول
                values = [
                    str(idx + 1),  # العدد التسلسلي للدورة
                    course_name,  # اسم الدورة
                    start_date_str,  # تاريخ بداية الدورة
                    end_date_str,  # تاريخ نهاية الدورة
                    str(total_course_students),  # القوة (إجمالي عدد الطلاب في الدورة)
                    str(not_started_count),  # عدد حالات "لم يباشر"
                    str(absent_count),  # عدد حالات الغياب
                    str(late_count),  # عدد حالات التأخير
                    str(excused_count)  # عدد حالات الغياب بعذر
                ]

                # إضافة القيم للحالات الموجودة فقط
                if show_field_app:
                    values.append(str(field_app_count))
                if show_student_day:
                    values.append(str(student_day_count))
                if show_evening_remote:
                    values.append(str(evening_remote_count))
                if show_death_case:
                    values.append(str(death_case_count))
                if show_hospital:
                    values.append(str(hospital_count))

                # إضافة العدد الفعلي دائماً
                values.append(str(effective_count))

                # إدخال القيم في الجدول بالترتيب من اليمين إلى اليسار
                for i, value in enumerate(values):
                    idx = len(values) - i - 1
                    row_cells[idx].text = value

                # تنسيق الخلايا
                for cell in row_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in cell.paragraphs[0].runs:
                        run.font.rtl = True
                        run.font.size = Pt(10)

            # إضافة فقرة فاصلة قبل التوقيعات
            doc.add_paragraph()

            # إضافة جدول للتوقيعات
            signatures_table = doc.add_table(rows=1, cols=3)
            signatures_table.style = 'Table Grid'

            sig_cells = signatures_table.rows[0].cells
            sig_cells[2].text = "اسم المراقب: ____________"
            sig_cells[1].text = "رئيس قسم الفصول: ______________"
            sig_cells[0].text = "مدير قسم شؤون المدربين: _____________"

            for cell in sig_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell.paragraphs[0].runs:
                    run.font.rtl = True
                    run.font.size = Pt(11)

            # ============================================================================
            # إضافة صفحة جديدة للتفاصيل - فقط للحالات المطلوبة
            # ============================================================================

            # دالة مساعدة لإضافة جدول الطلاب
            def add_students_table(title, students_data, has_reason=False):
                """إضافة جدول الطلاب لحالة معينة"""
                if not students_data:
                    return  # تخطي إذا لم تكن هناك بيانات

                # إضافة صفحة جديدة
                doc.add_page_break()

                # إضافة عنوان الصفحة
                title_para = doc.add_heading(f'{title} ليوم {arabic_weekday} بتاريخ {arabic_date}', level=1)
                title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in title_para.runs:
                    run.font.rtl = True
                    run.font.bold = True
                    run.font.size = Pt(16)

                # إضافة فاصل
                border_para = doc.add_paragraph()
                border_para.paragraph_format.border_bottom = True

                # إنشاء جدول الطلاب
                cols_count = 6 if has_reason else 5  # إضافة عمود للسبب عند الحاجة
                students_table = doc.add_table(rows=1, cols=cols_count)
                students_table.style = 'Table Grid'

                # تحديد عناوين الجدول
                headers = ["العدد", "الاسم", "الرتبة", "رقم الهوية", "الدورة"]
                if has_reason:
                    headers.append("السبب")  # إضافة عمود السبب إذا كان مطلوباً

                header_cells = students_table.rows[0].cells
                for i, header in enumerate(headers):
                    idx = len(headers) - i - 1  # لترتيب العناوين من اليمين لليسار
                    header_cells[idx].text = header
                    header_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    for run in header_cells[idx].paragraphs[0].runs:
                        run.font.bold = True
                        run.font.rtl = True
                        run.font.size = Pt(12)

                    # إضافة تظليل للعناوين
                    try:
                        shading_elm = parse_xml(r'<w:shd {} w:fill="DDDDDD"/>'.format(nsdecls('w')))
                        header_cells[idx]._element.get_or_add_tcPr().append(shading_elm)
                    except:
                        pass

                # إضافة بيانات الطلاب
                for idx, student in enumerate(students_data):
                    national_id, name, rank, course, status, reason = student

                    row_cells = students_table.add_row().cells

                    if has_reason:
                        row_cells[5].text = str(idx + 1)  # العدد التسلسلي
                        row_cells[4].text = name  # الاسم
                        row_cells[3].text = rank  # الرتبة
                        row_cells[2].text = national_id  # رقم الهوية
                        row_cells[1].text = course  # اسم الدورة
                        row_cells[0].text = reason if reason else "لم يحدد سبب"  # السبب
                    else:
                        row_cells[4].text = str(idx + 1)  # العدد التسلسلي
                        row_cells[3].text = name  # الاسم
                        row_cells[2].text = rank  # الرتبة
                        row_cells[1].text = national_id  # رقم الهوية
                        row_cells[0].text = course  # اسم الدورة

                    # تنسيق خلايا البيانات
                    for cell in row_cells:
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in cell.paragraphs[0].runs:
                            run.font.rtl = True
                            run.font.size = Pt(11)

                # إضافة إجمالي عدد الطلاب
                summary_para = doc.add_paragraph()
                summary_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                summary_text = summary_para.add_run(f"إجمالي عدد الطلاب: {len(students_data)}")
                summary_text.font.bold = True
                summary_text.font.rtl = True
                summary_text.font.size = Pt(12)

                # إضافة جدول التوقيعات
                doc.add_paragraph()
                signatures_table = doc.add_table(rows=1, cols=3)
                signatures_table.style = 'Table Grid'

                sig_cells = signatures_table.rows[0].cells
                sig_cells[2].text = "اسم المراقب: ____________"
                sig_cells[1].text = "رئيس قسم الفصول: ______________"
                sig_cells[0].text = "مدير قسم شؤون المدربين: _____________"

                for cell in sig_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in cell.paragraphs[0].runs:
                        run.font.rtl = True
                        run.font.size = Pt(11)

            # إضافة جداول الطلاب لكل حالة (فقط للحالات المطلوبة وتجاهل الحالات المذكورة)
            # إضافة جدول الطلاب الغائبين
            if absent_data:
                add_students_table("بيان الطلاب الغائبين", absent_data)

            # إضافة جدول الطلاب الغائبين بعذر
            if excused_data:
                add_students_table("بيان الطلاب الغائبين بعذر", excused_data, has_reason=True)

            # إضافة جدول الطلاب المتأخرين
            if late_data:
                add_students_table("بيان الطلاب المتأخرين", late_data)

            # إضافة جدول الطلاب في حالة "لم يباشر"
            if not_started_data:
                add_students_table("بيان الطلاب في حالة لم يباشر", not_started_data)

            # لا نضيف الجداول التالية كما هو مطلوب:
            # - بيان الطلاب في التطبيق الميداني
            # - بيان طلاب يوم الطالب
            # - بيان طلاب المسائية/عن بعد

            # نضيف فقط جداول حالات الوفاة والمنومين
            if death_case_data:
                add_students_table("بيان الطلاب في حالة وفاة", death_case_data, has_reason=True)

            # إضافة جدول الطلاب المنومين
            if hospital_data:
                add_students_table("بيان الطلاب المنومين", hospital_data, has_reason=True)

            # حفظ المستند
            export_file = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile=f"تكميل_الدورات_{selected_date_str}.docx"
            )

            if export_file:
                doc.save(export_file)
                messagebox.showinfo("نجاح", f"تم تصدير تكميل الدورات بنجاح إلى:\n{export_file}")
                # محاولة فتح الملف تلقائيًا
                try:
                    os.startfile(export_file)
                except:
                    pass

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير البيانات: {str(e)}")

    def update_statistics(self):
        cursor = self.conn.cursor()
        # احتساب عدد الطلاب غير المستبعدين فقط
        cursor.execute("SELECT COUNT(*) FROM trainees WHERE is_excluded=0")
        total_students = cursor.fetchone()[0]
        self.total_students_var.set(str(total_students))

        date_str = self.date_entry.get_date().strftime("%Y-%m-%d")

        # استعلام الحضور للطلاب غير المستبعدين فقط
        cursor.execute("""
            SELECT a.status, COUNT(*) 
            FROM attendance a
            JOIN trainees t ON a.national_id = t.national_id
            WHERE a.date=? AND t.is_excluded=0
            GROUP BY a.status
        """, (date_str,))

        status_counts = {row[0]: row[1] for row in cursor.fetchall()}

        present_count = status_counts.get("حاضر", 0)
        late_count = status_counts.get("متأخر", 0)
        excused_count = status_counts.get("غائب بعذر", 0)
        absent_count = status_counts.get("غائب", 0)
        not_started_count = status_counts.get("لم يباشر", 0)
        field_application_count = status_counts.get("تطبيق ميداني", 0)
        student_day_count = status_counts.get("يوم طالب", 0)
        evening_remote_count = status_counts.get("مسائية / عن بعد", 0)

        self.present_students_var.set(str(present_count))
        self.late_students_var.set(str(late_count))
        self.excused_students_var.set(str(excused_count))
        self.absent_students_var.set(str(absent_count))
        self.not_started_students_var.set(str(not_started_count))
        self.field_application_var.set(str(field_application_count))
        self.student_day_var.set(str(student_day_count))
        self.evening_remote_var.set(str(evening_remote_count))

        death_case_count = status_counts.get("حالة وفاة", 0)
        hospital_count = status_counts.get("منوم", 0)

        self.death_case_var.set(str(death_case_count))
        self.hospital_var.set(str(hospital_count))

        if total_students > 0:
            # إضافة الحالات الجديدة لحساب نسبة الحضور
            attendance_rate = ((present_count + late_count + field_application_count +
                                student_day_count + evening_remote_count) / total_students) * 100
        else:
            attendance_rate = 0.0

        self.attendance_rate_var.set(f"{attendance_rate:.2f}%")

    def check_student_absence(self, national_id, current_date):
        """
        فحص حالة غياب الطالب ورصد الغياب المتكرر
        يتم استدعاء هذه الدالة عند تسجيل غياب جديد للطالب

        المدخلات:
            national_id: رقم هوية الطالب
            current_date: تاريخ اليوم الحالي بتنسيق YYYY-MM-DD

        المخرجات:
            tuple(bool, str): الأول يشير إلى وجود غياب متكرر، والثاني هو نص رسالة التنبيه
        """
        cursor = self.conn.cursor()

        # الحصول على معلومات الطالب
        cursor.execute("""
            SELECT name, rank, course 
            FROM trainees 
            WHERE national_id=?
        """, (national_id,))

        student_info = cursor.fetchone()
        if not student_info:
            return False, ""

        student_name, student_rank, student_course = student_info

        # تحويل التاريخ الحالي إلى كائن تاريخ
        current_date_obj = datetime.datetime.strptime(current_date, "%Y-%m-%d").date()

        # الحصول على تواريخ حضور الطالب في الأيام السابقة (دون اليوم الحالي)
        cursor.execute("""
            SELECT date, status 
            FROM attendance 
            WHERE national_id=? AND date < ? 
            ORDER BY date DESC
        """, (national_id, current_date))

        attendance_records = cursor.fetchall()

        # حساب عدد أيام الغياب المتتالية
        consecutive_absences = 0

        # حساب عدد أيام الغياب الإجمالية
        total_absences = 0

        # التحقق أولاً ما إذا كان اليوم المسجل هو "غائب"
        consecutive_absences = 1  # اليوم الحالي محسوب كغياب (لأننا ندعو هذه الدالة فقط عند تسجيل غياب)
        total_absences = 1

        # معالجة سجلات الحضور السابقة
        last_date = current_date_obj
        for record in attendance_records:
            date_str, status = record
            record_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()

            # حساب إجمالي الغياب (غائب وغائب بعذر يُحسبان كغياب للإجمالي)
            if status in ["غائب", "غائب بعذر"]:
                total_absences += 1

                # التحقق من التتابع - يجب أن يكون الفرق يوم واحد للاعتبار متتاليًا
                if (last_date - record_date).days == 1 and status == "غائب":
                    consecutive_absences += 1
                else:
                    # إذا كان هناك انقطاع في التتابع، نتوقف عن العد
                    if status != "غائب":  # إذا كان "غائب بعذر" لا يُحسب في التتابع
                        continue
            else:
                # إذا كان الطالب حاضرًا أو حالة أخرى، نتوقف عن عد الأيام المتتالية
                break

            last_date = record_date

        # تحديد نوع التنبيه المطلوب
        alert_message = ""
        show_alert = False

        if consecutive_absences >= 3:
            show_alert = True

            if consecutive_absences >= 4:
                # تنبيه أحمر للغياب المتكرر أكثر من 3 أيام متتالية
                alert_message = f"⚠️ تنبيه هام: الطالب {student_name} ({student_rank}) متغيب {consecutive_absences} أيام متتالية!\n\n" \
                                f"✓ الإجراء المطلوب: يجب اتخاذ إجراء بشأن هذا الطالب حسب التعليمات المنظمة.\n" \
                                f"✓ الدورة: {student_course}\n" \
                                f"✓ إجمالي أيام الغياب: {total_absences} أيام"
                alert_type = "خطير"
                alert_color = "red"
            else:
                # تنبيه أصفر للغياب 3 أيام متتالية
                alert_message = f"⚠️ تنبيه: الطالب {student_name} ({student_rank}) متغيب {consecutive_absences} أيام متتالية.\n\n" \
                                f"✓ الدورة: {student_course}\n" \
                                f"✓ إجمالي أيام الغياب: {total_absences} أيام"
                alert_type = "متوسط"
                alert_color = "orange"

        return show_alert, alert_message, alert_type if show_alert else None, alert_color if show_alert else None

    def insert_attendance_record(self, status, excuse_reason=""):
        """تعديل دالة تسجيل الحضور الموجودة لإضافة فحص الغياب المتكرر"""
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        national_id = self.id_entry.get().strip()
        if not national_id:
            messagebox.showwarning("تنبيه", "الرجاء اختيار طالب من خلال البحث بالاسم أو الهوية")
            return

        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT national_id, name, rank, course, is_excluded 
            FROM trainees 
            WHERE national_id=?
        """, (national_id,))

        trainee = cursor.fetchone()
        if not trainee:
            messagebox.showwarning("تنبيه", "لا يوجد طالب بهذا الرقم")
            return

        # التحقق من استبعاد الطالب
        if trainee[4] == 1:
            messagebox.showwarning("تنبيه", "هذا الطالب مستبعد ولا يمكن تسجيل حضوره")
            return

        current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?", (trainee[0], current_date))
        existing_record = cursor.fetchone()

        if existing_record:
            existing_status = existing_record[0]

            # استخدام نوافذ خطأ بدلاً من معلومات لجذب انتباه المستخدم
            if existing_status == status:
                # إذا كانت نفس الحالة
                messagebox.showerror("خطأ في التكرار",
                                     f"⚠️ تنبيه: الطالب {trainee[1]} مسجل بالفعل بحالة '{existing_status}' اليوم\n\nلا يمكن تكرار نفس الحالة للطالب في نفس اليوم.")
            else:
                # إذا كانت حالة مختلفة
                messagebox.showerror("تعارض في الحالة",
                                     f"⚠️ تنبـــيه: الطالب {trainee[1]} مسجل بالفعل بحالة '{existing_status}' اليوم\n\nلتغيير الحالة من '{existing_status}' إلى '{status}'، يرجى استخدام خاصية تعديل الحضور من قائمة سجل الحضور.")

            # مسح قيمة الهوية
            self.id_entry.delete(0, tk.END)
            self.name_search_entry.delete(0, tk.END)
            return

        # التحقق من حالة الطالب في اليوم السابق
        current_date_obj = self.date_entry.get_date()
        yesterday_date_obj = current_date_obj - datetime.timedelta(days=1)
        yesterday_date = yesterday_date_obj.strftime("%Y-%m-%d")

        cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?", (trainee[0], yesterday_date))
        yesterday_record = cursor.fetchone()

        # التحقق إذا كان الطالب مسجل "لم يباشر" في اليوم السابق ويحاول المستخدم تسجيله "غائب"
        if yesterday_record and yesterday_record[0] == "لم يباشر" and status == "غائب":
            response = messagebox.askquestion("تنبيه هام ⚠️",
                                              f"الطالب {trainee[1]} مسجل كـ 'لم يباشر' في اليوم السابق.\n\n"
                                              "• تأكد من أن الطالب لم يباشر الدورة فعلاً.\n"
                                              "• إذا حضر الطالب اليوم، يجب تسجيله كـ 'حاضر' أو 'متأخر'.\n"
                                              "• استمرار تسجيله كـ 'غائب' يعتبر مخالف لتعلميات التدريب المستديمة.\n\n"
                                              "هل تريد تغيير الحالة إلى 'لم يباشر' بدلاً من 'غائب'؟",
                                              icon="warning")
            if response == "yes":
                status = "لم يباشر"
            elif response == "no":
                # إضافة تأكيد إضافي عند الإصرار على الغياب
                confirm = messagebox.askquestion("تأكيد نهائي",
                                                 f"هل أنت متأكد من تسجيل الطالب {trainee[1]} كـ 'غائب' رغم أنه 'لم يباشر' بالأمس؟",
                                                 icon="warning")
                if confirm != "yes":
                    return

        t_id, t_name, t_rank, t_course, _ = trainee
        current_time = datetime.datetime.now().strftime("%H:%M:%S")

        # معالجة تسجيل حالة الغياب ورصد الغياب المتكرر
        absence_alert = False
        alert_message = ""
        alert_type = None
        alert_color = None

        if status in ["غائب"]:
            absence_alert, alert_message, alert_type, alert_color = self.check_student_absence(t_id, current_date)

        try:
            with self.conn:
                self.conn.execute("""
                    INSERT INTO attendance (
                        national_id, name, rank, course,
                        time, date, status, original_status,
                        registered_by, excuse_reason,
                        updated_by, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    t_id, t_name, t_rank, t_course,
                    current_time, current_date,
                    status, status,
                    self.current_user["full_name"], excuse_reason,
                    "", ""
                ))

            # تحديث رسالة التأكيد في عنصر الواجهة بدلاً من نافذة منبثقة
            if status == "حاضر":
                icon_status = "✅"
            elif status == "غائب":
                icon_status = "❌"
            elif status == "متأخر":
                icon_status = "⏰"
            elif status == "غائب بعذر":
                icon_status = "📝"
            elif status == "لم يباشر":
                icon_status = "⏳"
            else:
                icon_status = "📌"

            # نعرض الرسالة فقط في حقل آخر طالب سُجّل بدلاً من نافذة منبثقة
            self.last_registered_label.config(text=f"آخر طالب سُجِّل: {t_name} ({status}) {icon_status}")

            # مسح حقول الإدخال
            self.id_entry.delete(0, tk.END)
            self.name_search_entry.delete(0, tk.END)
            self.name_listbox.delete(0, tk.END)

            self.update_statistics()
            self.update_attendance_display()

            # عرض تنبيه الغياب المتكرر إذا كان مطلوبًا
            if absence_alert:
                self.show_absence_alert(alert_message, alert_type, alert_color)

        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def show_absence_alert(self, message, alert_type, alert_color):
        """عرض نافذة تنبيه مخصصة للغياب المتكرر"""
        alert_window = tk.Toplevel(self.root)
        alert_window.title(f"تنبيه غياب متكرر - {alert_type}")
        alert_window.geometry("650x500")  # نافذة أكبر
        alert_window.configure(bg="#FFFFFF")
        alert_window.transient(self.root)
        alert_window.grab_set()

        # توسيط النافذة
        x = (alert_window.winfo_screenwidth() - 650) // 2
        y = (alert_window.winfo_screenheight() - 500) // 2
        alert_window.geometry(f"650x500+{x}+{y}")

        # إطار العنوان
        title_frame = tk.Frame(alert_window, bg=alert_color, padx=10, pady=15)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame,
            text="⚠️ تنبيه غياب متكرر ⚠️",
            font=("Tajawal", 20, "bold"),  # خط أكبر وغامق
            bg=alert_color,
            fg="white"
        )
        title_label.pack()

        # إطار الرسالة
        message_frame = tk.Frame(alert_window, bg="#FFFFFF", padx=20, pady=20)
        message_frame.pack(fill=tk.BOTH, expand=True)

        message_text = tk.Text(
            message_frame,
            wrap=tk.WORD,
            bg="#FFFFFF",
            font=("Tajawal", 16, "bold"),  # خط أكبر وغامق
            relief=tk.FLAT,
            height=8
        )
        message_text.insert(tk.END, message)
        message_text.configure(state="disabled")  # جعل النص للقراءة فقط
        message_text.pack(fill=tk.BOTH, expand=True)

        # أزرار الإجراءات
        button_frame = tk.Frame(alert_window, bg="#FFFFFF", padx=20, pady=15)
        button_frame.pack(fill=tk.X)

        ok_button = tk.Button(
            button_frame,
            text="موافق",
            font=("Tajawal", 14, "bold"),  # خط أكبر
            bg="#4CAF50",
            fg="white",
            padx=25,  # زيادة حجم الزر
            pady=8,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=alert_window.destroy
        )
        ok_button.pack(side=tk.RIGHT, padx=10)

        # إضافة زر خاص إذا كان التنبيه من النوع الخطير
        if alert_type == "خطير":
            action_button = tk.Button(
                button_frame,
                text="اتخاذ إجراء",
                font=("Tajawal", 14, "bold"),  # خط أكبر
                bg="#FF5722",
                fg="white",
                padx=25,  # زيادة حجم الزر
                pady=8,
                bd=0,
                relief=tk.FLAT,
                cursor="hand2",
                command=lambda: self.take_absence_action(alert_window)
            )
            action_button.pack(side=tk.LEFT, padx=10)

    def take_absence_action(self, parent_window=None):
        """عرض نافذة اتخاذ إجراء للغياب المتكرر"""
        # نافذة بسيطة لعرض الإجراءات المطلوبة
        messagebox.showinfo(
            "إجراءات متابعة الغياب",
            "يرجى اتخاذ الإجراءات التالية:\n\n"
            "1. التواصل مع الطالب لمعرفة سبب الغياب\n"
            "2. إبلاغ المشرف المباشر عن حالة الغياب المتكرر\n"
            "3. توثيق حالة الغياب في سجل الطالب\n"
            "4. متابعة الحالة مع إدارة الدورة\n\n"
            "ملاحظة: يمكن استخدام نموذج متابعة الغياب المعتمد من القسم"
        )

        # إغلاق نافذة التنبيه إذا كانت مفتوحة
        if parent_window:
            parent_window.destroy()

    def process_barcode_ids(self, status):
        """تعديل دالة معالجة الباركود لإضافة فحص الغياب المتكرر"""
        if not self.current_user["permissions"]["can_edit_attendance"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تسجيل الحضور والغياب")
            return

        # قراءة النص من مربع الإدخال
        barcode_text = self.barcode_text.get(1.0, tk.END).strip()
        if not barcode_text:
            messagebox.showinfo("تنبيه", "الرجاء إدخال أرقام الهويات أولاً")
            return

        # تقسيم النص إلى أسطر للحصول على أرقام الهويات
        id_lines = [line.strip() for line in barcode_text.split("\n") if line.strip()]
        if not id_lines:
            messagebox.showinfo("تنبيه", "لم يتم العثور على أرقام هويات صالحة")
            return

        # الحصول على التاريخ الحالي ووقت التسجيل
        current_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        current_time = datetime.datetime.now().strftime("%H:%M:%S")

        cursor = self.conn.cursor()

        # قوائم لتتبع النتائج
        successful_ids = []
        failed_ids = []
        already_registered_ids = []
        excluded_ids = []
        absence_alerts = []  # لتخزين معلومات تنبيهات الغياب المتكرر

        # معالجة كل رقم هوية
        for national_id in id_lines:
            # تخطي القيم الفارغة
            if not national_id:
                continue

            try:
                # التحقق من وجود الطالب وما إذا كان مستبعدًا
                cursor.execute("""
                    SELECT national_id, name, rank, course, is_excluded 
                    FROM trainees 
                    WHERE national_id=?
                """, (national_id,))

                trainee = cursor.fetchone()
                if not trainee:
                    failed_ids.append(national_id)
                    continue

                # التحقق من استبعاد الطالب
                if trainee[4] == 1:
                    excluded_ids.append(national_id)
                    continue

                # التحقق مما إذا كان الطالب مسجلاً بالفعل لهذا اليوم
                cursor.execute("SELECT status FROM attendance WHERE national_id=? AND date=?",
                               (trainee[0], current_date))
                existing_record = cursor.fetchone()

                if existing_record:
                    already_registered_ids.append(national_id)
                    continue

                # فحص تنبيهات الغياب إذا كان التسجيل غيابًا
                if status == "غائب":
                    absence_alert, alert_message, alert_type, alert_color = self.check_student_absence(trainee[0],
                                                                                                       current_date)
                    if absence_alert:
                        absence_alerts.append((alert_message, alert_type, alert_color))

                # إدراج سجل حضور جديد
                with self.conn:
                    self.conn.execute("""
                        INSERT INTO attendance (
                            national_id, name, rank, course,
                            time, date, status, original_status,
                            registered_by, excuse_reason,
                            updated_by, updated_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        trainee[0], trainee[1], trainee[2], trainee[3],
                        current_time, current_date,
                        status, status,
                        self.current_user["full_name"], "",
                        "", ""
                    ))

                successful_ids.append(national_id)

            except Exception as e:
                print(f"خطأ في معالجة الهوية {national_id}: {str(e)}")
                failed_ids.append(national_id)

        # إعداد رسالة ملخص النتائج
        result_message = f"تمت معالجة {len(id_lines)} رقم هوية:\n\n"

        if successful_ids:
            result_message += f"✅ تم تسجيل {len(successful_ids)} طالب بنجاح بحالة '{status}'.\n"

        if already_registered_ids:
            result_message += f"⚠️ {len(already_registered_ids)} طالب مسجل مسبقاً في هذا اليوم.\n"

        if excluded_ids:
            result_message += f"❌ {len(excluded_ids)} طالب مستبعد لا يمكن تسجيل حضورهم.\n"

        if failed_ids:
            result_message += f"❓ {len(failed_ids)} رقم هوية غير موجود في قاعدة البيانات."

        # عرض النتائج
        messagebox.showinfo("نتائج تسجيل الحضور", result_message)

        # تفريغ مربع النص بعد المعالجة الناجحة إذا تم تسجيل طلاب بنجاح
        if successful_ids:
            self.barcode_text.delete(1.0, tk.END)

        # تحديث الإحصائيات وعرض الحضور
        self.update_statistics()
        self.update_attendance_display()

        # عرض تنبيهات الغياب المتكرر (إذا وجدت)
        if absence_alerts:
            # عرض التنبيه الأول فقط إذا كان هناك أكثر من تنبيه
            first_alert = absence_alerts[0]
            self.show_absence_alert(
                first_alert[0] + f"\n\nملاحظة: هناك {len(absence_alerts)} تنبيه غياب متكرر في هذه المجموعة."
                if len(absence_alerts) > 1 else first_alert[0],
                first_alert[1],
                first_alert[2]
            )

    def export_course_to_word(self, course_name):
        """وظيفة تصدير بيانات الدورة إلى ملف وورد مع جدول حضور فارغ للأيام بتنسيق عمودي"""
        if not self.current_user["permissions"]["can_export_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
            return

        try:
            # التأكد من وجود مكتبة python-docx
            if 'Document' not in globals():
                messagebox.showerror("خطأ",
                                     "لم يتم العثور على مكتبة python-docx. قم بتثبيتها باستخدام: pip install python-docx")
                return

            # الحصول على بيانات الطلاب في الدورة (فقط غير المستبعدين)
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT national_id, name, rank
                FROM trainees
                WHERE course=? AND is_excluded=0
                ORDER BY name
            """, (course_name,))
            students_data = cursor.fetchall()

            if not students_data:
                messagebox.showinfo("ملاحظة", f"لا يوجد طلاب نشطين مسجلين في الدورة '{course_name}'")
                return

            # إنشاء مستند جديد
            doc = Document()

            # إعداد المستند للغة العربية (RTL) بتنسيق عمودي
            section = doc.sections[0]
            section.page_width = Inches(8.27)  # A4 width in portrait
            section.page_height = Inches(11.69)  # A4 height in portrait
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            section.top_margin = Inches(0.7)
            section.bottom_margin = Inches(0.7)

            # إعداد الرأس (Header) مع خط فاصل
            header = section.header
            header_para = header.paragraphs[0]
            header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            header_run = header_para.add_run(f'كشف حضور وغياب طلاب دورة: {course_name}')
            header_run.font.size = Pt(14)
            header_run.font.bold = True
            header_run.font.rtl = True

            # إضافة إجمالي عدد الطلاب في الرأس
            header_para = header.add_paragraph()
            header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            student_count_run = header_para.add_run(f'إجمالي عدد الطلاب: {len(students_data)}')
            student_count_run.font.size = Pt(12)
            student_count_run.font.bold = True
            student_count_run.font.rtl = True

            # إضافة خط أفقي بعد معلومات الدورة في الرأس
            header_para.paragraph_format.border_bottom = True

            # إضافة تاريخ الطباعة في الرأس
            today_date = datetime.datetime.now().strftime("%Y-%m-%d")
            header_para = header.add_paragraph()
            header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            header_date = header_para.add_run(f'تاريخ الطباعة: {today_date}')
            header_date.font.size = Pt(9)
            header_date.font.rtl = True

            # إعداد التذييل بشكل بسيط
            footer = section.footer
            footer_para = footer.paragraphs[0]
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            footer_text = footer_para.add_run('نظام إدارة الحضور والغياب - قسم شؤون المدربين')
            footer_text.font.size = Pt(9)
            footer_text.font.rtl = True

            # إضافة فقرة فاصلة قبل الجدول
            doc.add_paragraph()

            # إنشاء جدول للحضور والغياب
            table = doc.add_table(rows=1, cols=8)
            table.style = 'Table Grid'

            # تعريف رأس الجدول
            hdr_cells = table.rows[0].cells
            headers = ["العدد", "الاسم", "رقم الهوية", "الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس"]

            # إضافة العناوين من اليمين إلى اليسار (عكس الترتيب)
            for i, header in enumerate(reversed(headers)):
                hdr_cells[i].text = header
                # تنسيق العناوين
                hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in hdr_cells[i].paragraphs[0].runs:
                    run.font.bold = True
                    run.font.size = Pt(11)
                    run.font.rtl = True

                # تطبيق تظليل لرأس الجدول بطريقة بسيطة
                try:
                    shading_elm = parse_xml(r'<w:shd {} w:fill="D9D9D9"/>'.format(nsdecls('w')))
                    hdr_cells[i]._element.get_or_add_tcPr().append(shading_elm)
                except:
                    # في حالة حدوث خطأ، نتجاهل التظليل
                    pass

            # إضافة بيانات الطلاب
            for i, student in enumerate(students_data):
                national_id, name, rank = student
                row_cells = table.add_row().cells

                # إضافة البيانات من اليمين إلى اليسار (عكس الترتيب)
                # العدد (تسلسلي)
                row_cells[7].text = str(i + 1)
                row_cells[7].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # الاسم - تغيير المحاذاة إلى توسيط
                row_cells[6].text = name
                row_cells[6].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # رقم الهوية
                row_cells[5].text = national_id
                row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # الأيام تبقى فارغة للتعبئة يدوياً
                for day_idx in range(5):
                    row_cells[day_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # تنسيق النص في الصف
                for cell in row_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.rtl = True
                            run.font.size = Pt(10)

            # ضبط أبعاد الجدول لتناسب التنسيق العمودي - زيادة عرض عمود الاسم
            table.autofit = False
            col_widths = [0.5, 2.6, 1.4, 0.7, 0.7, 0.7, 0.7, 0.7]  # زيادة عرض عمود الاسم (2.6 بدلاً من 2.0)

            # تطبيق العرض المحدد لكل عمود
            try:
                for i, width in enumerate(col_widths):
                    table.columns[i].width = Inches(width)
            except:
                # في حالة حدوث خطأ، نتجاهل تعديل العرض
                pass

            # إضافة مساحة بعد الجدول
            doc.add_paragraph()

            # إضافة جدول للتوقيعات
            sig_table = doc.add_table(rows=1, cols=3)
            sig_table.style = 'Table Grid'
            sig_cells = sig_table.rows[0].cells

            sig_cells[2].text = "المسؤول: _________________"
            sig_cells[1].text = "رئيس القسم: ______________"
            sig_cells[0].text = "المدير: __________________"

            for cell in sig_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell.paragraphs[0].runs:
                    run.font.rtl = True
                    run.font.size = Pt(11)

            # إضافة ملاحظات في نهاية المستند
            doc.add_paragraph()
            notes_para = doc.add_paragraph()
            notes_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            notes_para.add_run("ملاحظات:").bold = True

            # إضافة خطوط للملاحظات
            for _ in range(3):
                line_para = doc.add_paragraph("_" * 80)
                line_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # حفظ المستند
            export_file = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile=f"كشف_حضور_{course_name}.docx"
            )

            if export_file:
                doc.save(export_file)
                messagebox.showinfo("نجاح", f"تم تصدير كشف الحضور للدورة '{course_name}' بنجاح إلى:\n{export_file}")
                # فتح الملف مباشرة بعد التصدير
                try:
                    os.startfile(export_file)
                except:
                    # في حالة عدم تمكن النظام من فتح الملف، تجاهل الخطأ
                    pass

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير بيانات الدورة: {str(e)}")

    def export_course_diligence_behavior(self, course_name):
        """وظيفة تصدير بيان المواظبة والسلوك للدورة بتنسيق Word مع ترتيب الطلاب حسب الدرجة"""
        if not self.current_user["permissions"]["can_export_data"]:
            messagebox.showwarning("تنبيه", "ليس لديك صلاحية تصدير البيانات")
            return

        try:
            # التأكد من وجود مكتبة python-docx
            if 'Document' not in globals():
                messagebox.showerror("خطأ",
                                     "لم يتم العثور على مكتبة python-docx. قم بتثبيتها باستخدام: pip install python-docx")
                return

            # الحصول على بيانات الطلاب في الدورة (فقط غير المستبعدين)
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT national_id, name, rank
                FROM trainees
                WHERE course=? AND is_excluded=0
                ORDER BY name
            """, (course_name,))
            students_data = cursor.fetchall()

            if not students_data:
                messagebox.showinfo("ملاحظة", f"لا يوجد طلاب نشطين مسجلين في الدورة '{course_name}'")
                return

            # إنشاء نافذة حالة لإظهار تقدم التصدير
            progress_window = tk.Toplevel(self.root)
            progress_window.title("جاري حساب المواظبة والسلوك")
            progress_window.geometry("400x150")
            progress_window.configure(bg=self.colors["light"])
            progress_window.transient(self.root)
            progress_window.resizable(False, False)
            progress_window.grab_set()

            x = (progress_window.winfo_screenwidth() - 400) // 2
            y = (progress_window.winfo_screenheight() - 150) // 2
            progress_window.geometry(f"400x150+{x}+{y}")

            tk.Label(
                progress_window,
                text=f"جاري حساب نتائج المواظبة والسلوك لدورة: {course_name}",
                font=self.fonts["text_bold"],
                bg=self.colors["light"],
                pady=10
            ).pack()

            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(
                progress_window,
                variable=progress_var,
                maximum=100,
                length=350
            )
            progress_bar.pack(pady=10)

            status_label = tk.Label(
                progress_window,
                text="جاري تحليل بيانات الحضور والغياب...",
                font=self.fonts["text"],
                bg=self.colors["light"]
            )
            status_label.pack(pady=5)

            progress_window.update()

            # إنشاء مستند جديد
            doc = Document()

            # إعداد المستند للغة العربية (RTL)
            section = doc.sections[0]
            section.page_width = Inches(8.27)  # A4 width
            section.page_height = Inches(11.69)  # A4 height
            section.left_margin = Inches(0.7)
            section.right_margin = Inches(0.7)
            section.top_margin = Inches(0.7)
            section.bottom_margin = Inches(0.7)

            # إضافة عنوان المستند
            title = doc.add_heading(f'بيان المواظبة والسلوك لطلاب دورة: {course_name}', level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in title.runs:
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.rtl = True

            # إضافة معلومات الطباعة والتاريخ
            date_info = doc.add_paragraph()
            date_info.alignment = WD_ALIGN_PARAGRAPH.LEFT
            today_date = datetime.datetime.now().strftime("%Y-%m-%d")
            date_run = date_info.add_run(f'تاريخ الطباعة: {today_date}')
            date_run.font.size = Pt(10)
            date_run.font.rtl = True

            # إضافة خط أفقي
            border_paragraph = doc.add_paragraph()
            border_paragraph.paragraph_format.border_bottom = True

            # إنشاء جدول للمواظبة والسلوك
            table = doc.add_table(rows=1, cols=6)
            table.style = 'Table Grid'

            # عناوين الجدول (من اليمين إلى اليسار)
            hdr_cells = table.rows[0].cells
            headers = ["عدد", "الاسم", "الرتبة", "رقم الهوية", "المواظبة", "السلوك"]

            for i, header in enumerate(headers):
                # حساب الموقع المناسب للعناوين (من اليمين إلى اليسار)
                idx = len(headers) - i - 1
                hdr_cells[idx].text = header
                hdr_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in hdr_cells[idx].paragraphs[0].runs:
                    run.font.bold = True
                    run.font.size = Pt(12)
                    run.font.rtl = True

                # تطبيق تظليل للرأس
                try:
                    shading_elm = parse_xml(r'<w:shd {} w:fill="D9D9D9"/>'.format(nsdecls('w')))
                    hdr_cells[idx]._element.get_or_add_tcPr().append(shading_elm)
                except:
                    pass

            # معالجة بيانات كل طالب وحساب درجة المواظبة
            student_scores = []
            total_students = len(students_data)

            for index, student in enumerate(students_data):
                national_id, name, rank = student

                # تحديث شريط التقدم
                progress_var.set((index / total_students) * 80)  # 80% للمعالجة
                status_label.config(text=f"معالجة الطالب {index + 1} من {total_students}: {name}")
                progress_window.update()

                # حساب درجة المواظبة:
                # 1. الدرجة الأولية هي 100
                # 2. خصم 4 درجات لكل غياب كامل
                # 3. خصم 1 درجة لكل تأخير
                # 4. خصم 0.5 درجة لكل غياب بعذر

                # الاستعلام عن حالات الحضور للطالب
                cursor.execute("""
                    SELECT status
                    FROM attendance
                    WHERE national_id=?
                """, (national_id,))
                attendance_records = cursor.fetchall()

                diligence_score = 100.0  # البداية من 100

                for record in attendance_records:
                    status = record[0]
                    if status == "غائب" or status == "غائب بعذر":  # تعديل: خصم 4 نقاط لـ "غائب بعذر"
                        diligence_score -= 4.0
                    elif status == "متأخر":
                        diligence_score -= 1.0
                    elif status == "حالة وفاة" or status == "منوم":  # إضافة: خصم 0.5 نقطة للحالات الجديدة
                        diligence_score -= 0.5

                # التأكد من عدم نزول الدرجة عن صفر
                diligence_score = max(0, diligence_score)

                # حفظ بيانات الطالب مع الدرجة
                student_scores.append((national_id, name, rank, diligence_score))

            # ترتيب الطلاب تصاعدياً حسب درجة المواظبة (الأقل يأتي أولاً)
            student_scores.sort(key=lambda x: x[3])

            # إضافة بيانات الطلاب إلى الجدول بعد الترتيب
            for index, (national_id, name, rank, diligence_score) in enumerate(student_scores):
                # تحديث شريط التقدم
                progress_var.set(80 + (index / total_students) * 15)  # 15% للترتيب والإضافة
                status_label.config(text=f"إضافة الطالب {index + 1} من {total_students} إلى التقرير")
                progress_window.update()

                # درجة السلوك دائمًا 100
                behavior_score = 100.0

                # إضافة صف جديد للطالب
                row_cells = table.add_row().cells

                # الترتيب من اليمين إلى اليسار
                row_cells[5].text = str(index + 1)  # العدد التسلسلي
                row_cells[4].text = name  # الاسم
                row_cells[3].text = rank  # الرتبة
                row_cells[2].text = national_id  # رقم الهوية
                row_cells[1].text = f"{diligence_score:.1f}"  # المواظبة بدقة رقم عشري واحد
                row_cells[0].text = f"{behavior_score:.0f}"  # السلوك (دائمًا 100)

                # تنسيق الخلايا
                for cell in row_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in cell.paragraphs[0].runs:
                        run.font.rtl = True
                        run.font.size = Pt(11)

                # تلوين الصف حسب درجة المواظبة
                if diligence_score < 90:  # إذا كانت الدرجة أقل من 90، تمييزها بلون فاتح
                    try:
                        for cell in row_cells:
                            shading_elm = parse_xml(r'<w:shd {} w:fill="FFDDDD"/>'.format(nsdecls('w')))
                            cell._element.get_or_add_tcPr().append(shading_elm)
                    except:
                        pass

            # تنسيق الجدول
            table.autofit = False
            try:
                # تعيين عرض الأعمدة (العرض بالبوصة)
                widths = [0.8, 0.8, 1.2, 1.5, 2.5, 0.5]  # السلوك، المواظبة، الهوية، الرتبة، الاسم، العدد
                for i, width in enumerate(widths):
                    table.columns[i].width = Inches(width)
            except:
                pass

            # إضافة فقرة فاصلة بعد الجدول
            doc.add_paragraph()

            # إضافة جدول للتوقيعات
            signature_table = doc.add_table(rows=1, cols=3)
            signature_table.style = 'Table Grid'

            sig_cells = signature_table.rows[0].cells
            sig_cells[2].text = "مسؤول الحضور: _________________"
            sig_cells[1].text = "رئيس القسم: __________________"
            sig_cells[0].text = "مدير التدريب: ________________"

            for cell in sig_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell.paragraphs[0].runs:
                    run.font.rtl = True
                    run.font.size = Pt(11)

            # إضافة نص توضيحي في نهاية المستند
            doc.add_paragraph()
            note_para = doc.add_paragraph()
            note_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            note_run = note_para.add_run("ملاحظات حساب المواظبة:")
            note_run.font.bold = True
            note_run.font.rtl = True

            notes = [
                "- تبدأ درجة المواظبة من 100 درجة.",
                "- يتم خصم 4 درجات عن كل يوم غياب.",
                "- يتم خصم 4 درجات عن كل غياب بعذر.",
                "- يتم خصم 1 درجة عن كل حالة تأخير.",
                "- يتم خصم 0.5 درجة عن كل حالة وفاة.",
                "- يتم خصم 0.5 درجة عن كل حالة منوم.",
                "- درجة السلوك 100 درجة للجميع."
            ]

            for note in notes:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.add_run(note).font.rtl = True

            # تحديث شريط التقدم
            progress_var.set(95)
            status_label.config(text="فتح حوار حفظ الملف...")
            progress_window.update()

            # حفظ المستند
            export_file = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile=f"بيان_المواظبة_والسلوك_{course_name}.docx"
            )

            if export_file:
                progress_var.set(95)
                status_label.config(text="جاري حفظ الملف...")
                progress_window.update()

                doc.save(export_file)

                progress_var.set(100)
                status_label.config(text="تم تصدير البيان بنجاح!")
                progress_window.update()

                # إغلاق نافذة التقدم بعد ثانيتين
                progress_window.after(2000, progress_window.destroy)

                messagebox.showinfo("نجاح",
                                    f"تم تصدير بيان المواظبة والسلوك للدورة '{course_name}' بنجاح إلى:\n{export_file}")

                # محاولة فتح الملف تلقائيًا
                try:
                    os.startfile(export_file)
                except:
                    pass
            else:
                progress_window.destroy()

        except Exception as e:
            try:
                progress_window.destroy()
            except:
                pass
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير بيان المواظبة والسلوك: {str(e)}")

    def manage_multi_section_courses(self):
        """وظيفة إدارة الدورات متعددة الفصول"""
        # إنشاء نافذة جديدة
        multi_window = tk.Toplevel(self.root)
        multi_window.title("إدارة الفصول و تصدير الكشوفات")
        multi_window.geometry("900x600")
        multi_window.configure(bg=self.colors["light"])
        multi_window.grab_set()
        multi_window.resizable(True, True)

        # توسيط النافذة
        x = (multi_window.winfo_screenwidth() - 900) // 2
        y = (multi_window.winfo_screenheight() - 600) // 2
        multi_window.geometry(f"900x600+{x}+{y}")

        # عنوان النافذة
        tk.Label(
            multi_window,
            text="إدارة الفصول و تصدير الكشوفات",
            font=self.fonts["large_title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        # تعديل: إضافة إطار جديد لعرض معلومات إجمالي الطلاب تحت العنوان مباشرة
        students_info_frame = tk.Frame(multi_window, bg=self.colors["light"], padx=10, pady=5)
        students_info_frame.pack(fill=tk.X)

        # تعديل: متغير لعرض إجمالي عدد الطلاب
        total_students_var = tk.StringVar(value="")
        total_students_label = tk.Label(
            students_info_frame,
            textvariable=total_students_var,
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["primary"]
        )
        total_students_label.pack(pady=5)

        # إطار اختيار الدورة
        course_frame = tk.Frame(multi_window, bg=self.colors["light"], padx=10, pady=10)
        course_frame.pack(fill=tk.X)

        tk.Label(
            course_frame,
            text="اختيار الدورة:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(side=tk.RIGHT, padx=5)

        # الحصول على قائمة الدورات
        cursor = self.conn.cursor()
        cursor.execute("SELECT DISTINCT course_name FROM course_sections")
        multi_section_courses = [row[0] for row in cursor.fetchall()]

        course_var = tk.StringVar()
        course_dropdown = ttk.Combobox(
            course_frame,
            textvariable=course_var,
            values=multi_section_courses,
            width=30,
            font=self.fonts["text"]
        )
        course_dropdown.pack(side=tk.RIGHT, padx=5)

        # زر تحديث قائمة الدورات
        refresh_btn = tk.Button(
            course_frame,
            text="تحديث",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=10, pady=2,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: update_sections_list()
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)

        #زر تعديل تاريخ الدورة
        edit_course_info_btn = tk.Button(
            course_frame,
            text="تعديل تواريخ الدورة",
            font=self.fonts["text_bold"],
            bg=self.colors["warning"],
            fg="white",
            padx=10, pady=2,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self.edit_course_dates(course_var.get())
        )
        edit_course_info_btn.pack(side=tk.LEFT, padx=5)

        # إضافة زر حذف الدورة كاملة (للمشرفين فقط)
        if self.current_user["permissions"]["is_admin"]:
            delete_course_btn = tk.Button(
                course_frame,
                text="حذف الدورة كاملة",
                font=self.fonts["text_bold"],
                bg=self.colors["danger"],
                fg="white",
                padx=10, pady=2,
                bd=0, relief=tk.FLAT,
                cursor="hand2",
                command=lambda: delete_entire_course()
            )
            delete_course_btn.pack(side=tk.LEFT, padx=5)


        # إطار عرض الفصول
        sections_frame = tk.LabelFrame(
            multi_window,
            text="الفصول المتاحة",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=10, pady=10
        )
        sections_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # قائمة الفصول
        list_frame = tk.Frame(sections_frame, bg=self.colors["light"])
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        list_scroll = tk.Scrollbar(list_frame)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        sections_listbox = tk.Listbox(
            list_frame,
            font=self.fonts["text"],
            selectbackground=self.colors["primary"],
            selectforeground="white",
            yscrollcommand=list_scroll.set
        )
        sections_listbox.pack(fill=tk.BOTH, expand=True)
        list_scroll.config(command=sections_listbox.yview)

        # إطار التفاصيل
        details_frame = tk.Frame(sections_frame, bg=self.colors["light"], width=350)
        details_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=5, pady=5)

        # عنوان التفاصيل
        section_title_var = tk.StringVar(value="اختر فصلاً لعرض تفاصيله")
        section_title = tk.Label(
            details_frame,
            textvariable=section_title_var,
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["primary"]
        )
        section_title.pack(pady=(0, 10))

        # تعديل: نحتفظ بمتغير عدد الطلاب للاستخدام الداخلي دون عرضه في إطار التفاصيل
        students_count_var = tk.StringVar(value="")

        # أزرار الإجراءات
        actions_frame = tk.Frame(details_frame, bg=self.colors["light"], pady=10)
        actions_frame.pack(fill=tk.X)

        export_attendance_btn = tk.Button(
            actions_frame,
            text="تصدير كشف حضور",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: export_section_attendance_sheet()
        )
        export_attendance_btn.pack(fill=tk.X, pady=5)

        export_diligence_btn = tk.Button(
            actions_frame,
            text="تصدير كشف المواظبة والسلوك",
            font=self.fonts["text_bold"],
            bg="#8E44AD",
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: export_section_diligence()
        )
        export_diligence_btn.pack(fill=tk.X, pady=5)

        view_students_btn = tk.Button(
            actions_frame,
            text="عرض الطلاب وإدارة الفصول",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: manage_section_students()
        )
        view_students_btn.pack(fill=tk.X, pady=5)

        rename_section_btn = tk.Button(
            actions_frame,
            text="تغيير اسم الفصل",
            font=self.fonts["text_bold"],
            bg=self.colors["warning"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: rename_section()
        )
        rename_section_btn.pack(fill=tk.X, pady=5)

        # إضافة زر حذف الفصل مع ترحيل الطلاب (متاح للجميع)
        delete_section_btn = tk.Button(
            actions_frame,
            text="حذف الفصل مع ترحيل الطلاب",
            font=self.fonts["text_bold"],
            bg=self.colors["danger"],
            fg="white",
            padx=10, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: delete_section_with_transfer()
        )
        delete_section_btn.pack(fill=tk.X, pady=5)

        # الإطار السفلي للأزرار
        bottom_frame = tk.Frame(multi_window, bg=self.colors["light"], pady=10)
        bottom_frame.pack(fill=tk.X, padx=10)

        add_section_btn = tk.Button(
            bottom_frame,
            text="إضافة فصل جديد",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: add_new_section()
        )
        add_section_btn.pack(side=tk.LEFT, padx=5)

        # هنا يتم إضافة الزر الجديد
        import_sections_btn = tk.Button(
            bottom_frame,
            text="استيراد تحديثات الفصول",
            font=self.fonts["text_bold"],
            bg="#FF9800",  # لون برتقالي للتمييز
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: import_section_updates()
        )
        import_sections_btn.pack(side=tk.LEFT, padx=5)

        # تعريف دالة للإغلاق مع تحديث البيانات
        def on_close_multi_window():
            multi_window.destroy()
            self.update_statistics()
            self.update_students_tree()
            self.update_attendance_display()  # إضافة هذا السطر لتحديث عرض سجل الحضور أيضاً

        close_btn = tk.Button(
            bottom_frame,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=on_close_multi_window  # استخدام الدالة الجديدة بدلاً من multi_window.destroy
        )
        close_btn.pack(side=tk.RIGHT, padx=5)

        def import_section_updates():
            """استيراد تحديثات توزيع الطلاب على الفصول من ملف Excel مع دعم الأعمدة باللغة العربية"""
            selected_course = course_var.get().strip()
            if not selected_course:
                messagebox.showwarning("تنبيه", "الرجاء اختيار دورة أولاً")
                return

            # اختيار ملف Excel
            file_path = filedialog.askopenfilename(
                title="اختر ملف تحديثات الفصول",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )

            if not file_path:
                return

            # إنشاء نافذة تقدم العملية
            progress_window = tk.Toplevel(multi_window)
            progress_window.title("استيراد تحديثات الفصول")
            progress_window.geometry("450x180")
            progress_window.configure(bg=self.colors["light"])
            progress_window.transient(multi_window)
            progress_window.grab_set()

            # توسيط النافذة
            x = (progress_window.winfo_screenwidth() - 450) // 2
            y = (progress_window.winfo_screenheight() - 180) // 2
            progress_window.geometry(f"450x180+{x}+{y}")

            tk.Label(
                progress_window,
                text=f"جاري معالجة تحديثات الفصول لدورة: {selected_course}",
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
                text="جاري قراءة ملف Excel...",
                font=self.fonts["text"],
                bg=self.colors["light"]
            )
            status_label.pack(pady=5)

            progress_window.update()

            try:
                # قراءة ملف Excel
                progress_var.set(10)
                status_label.config(text="جاري قراءة ملف Excel...")
                progress_window.update()

                df = pd.read_excel(file_path)

                # تعريف ترجمة أسماء الأعمدة (دعم العربية والإنجليزية)
                column_mapping = {
                    'رقم الهوية': 'national_id',
                    'الفصل': 'section_name',
                    'اسم الفصل': 'section_name',
                    'national_id': 'national_id',
                    'section_name': 'section_name'
                }

                # تحويل أسماء الأعمدة من العربية إلى الإنجليزية
                rename_dict = {}
                for orig_col in df.columns:
                    if orig_col in column_mapping:
                        rename_dict[orig_col] = column_mapping[orig_col]

                if rename_dict:
                    df = df.rename(columns=rename_dict)

                # التحقق من وجود الأعمدة المطلوبة
                has_id = any(col in ["رقم الهوية", "national_id"] for col in df.columns)
                has_section = any(col in ["الفصل", "اسم الفصل", "section_name"] for col in df.columns)

                if not (has_id and has_section):
                    progress_window.destroy()
                    messagebox.showwarning("تنبيه",
                                           f"يجب أن يحتوي ملف التحديثات على الأعمدة التالية:\n\n" +
                                           "• رقم الهوية (national_id)\n" +
                                           "• الفصل (section_name)")
                    return

                # التحقق من صحة الفصول المذكورة في الملف
                progress_var.set(20)
                status_label.config(text="التحقق من صحة بيانات الفصول...")
                progress_window.update()

                # الحصول على قائمة الفصول المتاحة في الدورة
                cursor = self.conn.cursor()
                cursor.execute("""
                    SELECT section_name
                    FROM course_sections
                    WHERE course_name=?
                """, (selected_course,))

                available_sections = {row[0] for row in cursor.fetchall()}

                # التحقق من وجود الفصول المذكورة في الملف
                unique_sections = set(df['section_name'].dropna())
                invalid_sections = unique_sections - available_sections

                if invalid_sections:
                    progress_window.destroy()
                    messagebox.showwarning("تنبيه",
                                           f"توجد فصول غير موجودة في الدورة: {', '.join(invalid_sections)}\n\n" +
                                           "الفصول المتاحة هي: " + ', '.join(available_sections))
                    return

                # التحقق من وجود الطلاب المذكورين في ملف التحديثات
                progress_var.set(30)
                status_label.config(text="التحقق من بيانات الطلاب...")
                progress_window.update()

                # الحصول على قائمة الطلاب في الدورة الحالية
                cursor.execute("""
                    SELECT national_id, name
                    FROM trainees
                    WHERE course=? AND is_excluded=0
                """, (selected_course,))

                students_dict = {row[0]: row[1] for row in cursor.fetchall()}

                # التحقق من وجود جميع الطلاب المذكورين في الملف
                student_ids = df['national_id'].astype(str).tolist()
                invalid_students = [sid for sid in student_ids if sid not in students_dict]

                if invalid_students:
                    if len(invalid_students) > 5:
                        invalid_display = ', '.join(invalid_students[:5]) + f' وغيرهم ({len(invalid_students) - 5})'
                    else:
                        invalid_display = ', '.join(invalid_students)

                    proceed = messagebox.askyesno("تنبيه - طلاب غير موجودين",
                                                  f"هناك {len(invalid_students)} طالب غير موجود في الدورة: {invalid_display}\n\n" +
                                                  "هل تريد المتابعة وتجاهل هؤلاء الطلاب؟")

                    if not proceed:
                        progress_window.destroy()
                        return

                # تحضير التغييرات
                progress_var.set(50)
                status_label.config(text="تحضير التغييرات...")
                progress_window.update()

                # الحصول على التوزيع الحالي للطلاب على الفصول
                cursor.execute("""
                    SELECT national_id, section_name
                    FROM student_sections
                    WHERE course_name=?
                """, (selected_course,))

                current_assignments = {row[0]: row[1] for row in cursor.fetchall()}

                # تحضير قائمة التغييرات
                changes = []
                no_changes = []
                new_assignments = []

                for _, row in df.iterrows():
                    student_id = str(row['national_id']).strip()
                    new_section = str(row['section_name']).strip()

                    # تخطي الطلاب غير الموجودين
                    if student_id not in students_dict:
                        continue

                    # التحقق إذا كان الطالب في فصل مختلف حاليًا
                    if student_id in current_assignments:
                        current_section = current_assignments[student_id]

                        if current_section != new_section:
                            # تغيير الفصل
                            changes.append((student_id, students_dict[student_id], current_section, new_section))
                        else:
                            # لا تغيير
                            no_changes.append((student_id, students_dict[student_id], current_section))
                    else:
                        # طالب جديد ليس في أي فصل سابقًا
                        new_assignments.append((student_id, students_dict[student_id], new_section))

                # عرض ملخص التغييرات المقترحة
                progress_var.set(70)
                status_label.config(text="تجهيز ملخص التغييرات...")
                progress_window.update()

                summary = f"ملخص التغييرات:\n\n"
                summary += f"• عدد الطلاب الذين سيتم نقلهم بين الفصول: {len(changes)}\n"
                summary += f"• عدد الطلاب الجدد المراد تسجيلهم في فصول: {len(new_assignments)}\n"
                summary += f"• عدد الطلاب بدون تغيير: {len(no_changes)}\n"

                if invalid_students:
                    summary += f"• عدد الطلاب غير الموجودين في الدورة: {len(invalid_students)} (سيتم تجاهلهم)\n"

                progress_window.destroy()

                # عرض نافذة ملخص التغييرات
                summary_window = tk.Toplevel(multi_window)
                summary_window.title("ملخص التغييرات المقترحة")
                summary_window.geometry("600x500")
                summary_window.configure(bg=self.colors["light"])
                summary_window.transient(multi_window)
                summary_window.grab_set()

                # توسيط النافذة
                x = (summary_window.winfo_screenwidth() - 600) // 2
                y = (summary_window.winfo_screenheight() - 500) // 2
                summary_window.geometry(f"600x500+{x}+{y}")

                tk.Label(
                    summary_window,
                    text="ملخص التغييرات المقترحة",
                    font=self.fonts["title"],
                    bg=self.colors["primary"],
                    fg="white",
                    padx=10, pady=10
                ).pack(fill=tk.X)

                # عرض ملخص التغييرات
                summary_frame = tk.Frame(summary_window, bg=self.colors["light"], padx=10, pady=10)
                summary_frame.pack(fill=tk.BOTH, expand=True)

                tk.Label(
                    summary_frame,
                    text=summary,
                    font=self.fonts["text"],
                    bg=self.colors["light"],
                    justify=tk.RIGHT,
                    anchor=tk.E
                ).pack(fill=tk.X, pady=10)

                # إنشاء نوتبوك لعرض التفاصيل
                details_notebook = ttk.Notebook(summary_frame)
                details_notebook.pack(fill=tk.BOTH, expand=True, pady=10)

                # تبويب الطلاب المنقولين
                if changes:
                    changes_frame = tk.Frame(details_notebook, bg=self.colors["light"])
                    details_notebook.add(changes_frame, text=f"طلاب سيتم نقلهم ({len(changes)})")

                    changes_list = tk.Text(changes_frame, font=self.fonts["text"], width=70, height=15)
                    changes_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

                    for student_id, name, old_section, new_section in changes:
                        changes_list.insert(tk.END,
                                            f"{name} ({student_id}): من فصل {old_section} إلى فصل {new_section}\n")

                    changes_list.configure(state="disabled")

                # تبويب الطلاب الجدد
                if new_assignments:
                    new_frame = tk.Frame(details_notebook, bg=self.colors["light"])
                    details_notebook.add(new_frame, text=f"تسجيلات جديدة ({len(new_assignments)})")

                    new_list = tk.Text(new_frame, font=self.fonts["text"], width=70, height=15)
                    new_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

                    for student_id, name, section in new_assignments:
                        new_list.insert(tk.END, f"{name} ({student_id}): تسجيل في فصل {section}\n")

                    new_list.configure(state="disabled")

                # أزرار التأكيد أو الإلغاء
                button_frame = tk.Frame(summary_window, bg=self.colors["light"], pady=10)
                button_frame.pack(fill=tk.X, padx=10)

                def apply_changes():
                    # تنفيذ التغييرات
                    try:
                        current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        progress_window = tk.Toplevel(summary_window)
                        progress_window.title("تنفيذ التغييرات")
                        progress_window.geometry("450x180")
                        progress_window.configure(bg=self.colors["light"])
                        progress_window.transient(summary_window)
                        progress_window.grab_set()

                        # توسيط النافذة
                        x = (progress_window.winfo_screenwidth() - 450) // 2
                        y = (progress_window.winfo_screenheight() - 180) // 2
                        progress_window.geometry(f"450x180+{x}+{y}")

                        tk.Label(
                            progress_window,
                            text="جاري تنفيذ التغييرات...",
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
                            text="جاري الإعداد...",
                            font=self.fonts["text"],
                            bg=self.colors["light"]
                        )
                        status_label.pack(pady=5)

                        progress_window.update()

                        with self.conn:
                            # تنفيذ التغييرات
                            total_operations = len(changes) + len(new_assignments)
                            operations_done = 0

                            # 1. تحديث الطلاب الذين سيتم نقلهم
                            if changes:
                                status_label.config(text="جاري تعديل تسجيلات الفصول الحالية...")
                                progress_window.update()

                                for student_id, _, _, new_section in changes:
                                    self.conn.execute("""
                                        UPDATE student_sections
                                        SET section_name=?, assigned_date=?
                                        WHERE national_id=? AND course_name=?
                                    """, (new_section, current_date, student_id, selected_course))

                                    operations_done += 1
                                    progress_percent = (operations_done / total_operations) * 100
                                    progress_var.set(progress_percent)
                                    progress_window.update()

                            # 2. إضافة الطلاب الجدد
                            if new_assignments:
                                status_label.config(text="جاري إضافة تسجيلات جديدة للفصول...")
                                progress_window.update()

                                for student_id, _, section in new_assignments:
                                    self.conn.execute("""
                                        INSERT OR REPLACE INTO student_sections
                                        (national_id, course_name, section_name, assigned_date)
                                        VALUES (?, ?, ?, ?)
                                    """, (student_id, selected_course, section, current_date))

                                    operations_done += 1
                                    progress_percent = (operations_done / total_operations) * 100
                                    progress_var.set(progress_percent)
                                    progress_window.update()

                        progress_var.set(100)
                        status_label.config(text="تم تنفيذ التغييرات بنجاح!")
                        progress_window.update()

                        # إغلاق نافذة التقدم بعد ثانيتين
                        progress_window.after(2000, progress_window.destroy)

                        # إغلاق نافذة الملخص
                        summary_window.destroy()

                        # عرض رسالة نجاح
                        messagebox.showinfo("نجاح", "تم تنفيذ تحديثات الفصول بنجاح!")

                        # تحديث القوائم
                        update_sections_list()

                    except Exception as e:
                        try:
                            progress_window.destroy()
                        except:
                            pass

                        messagebox.showerror("خطأ", f"حدث خطأ أثناء تنفيذ التغييرات: {str(e)}")

                confirm_btn = tk.Button(
                    button_frame,
                    text="تنفيذ التغييرات",
                    font=self.fonts["text_bold"],
                    bg=self.colors["success"],
                    fg="white",
                    padx=15, pady=5,
                    bd=0, relief=tk.FLAT,
                    cursor="hand2",
                    command=apply_changes
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
                    command=summary_window.destroy
                )
                cancel_btn.pack(side=tk.RIGHT, padx=5)

            except Exception as e:
                try:
                    progress_window.destroy()
                except:
                    pass

                messagebox.showerror("خطأ", f"حدث خطأ أثناء معالجة ملف التحديثات: {str(e)}")

        # الوظائف المساعدة ضمن النافذة
        def update_sections_list():
            """تحديث قائمة الفصول المتاحة للدورة المحددة"""
            selected_course = course_var.get().strip()
            sections_listbox.delete(0, tk.END)

            # تعديل: إعادة ضبط متغيرات العرض
            section_title_var.set("اختر فصلاً لعرض تفاصيله")
            students_count_var.set("")

            if not selected_course:
                total_students_var.set("")
                return

            # تعديل: تحديث إجمالي عدد الطلاب في الدورة
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT COUNT(DISTINCT t.national_id)
                FROM trainees t
                WHERE t.course=? AND t.is_excluded=0
            """, (selected_course,))

            total_count = cursor.fetchone()[0]
            total_students_var.set(f"إجمالي الطلاب الملتحقين بدورة \"{selected_course}\": {total_count}")

            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT section_name
                FROM course_sections
                WHERE course_name=?
                ORDER BY section_name
            """, (selected_course,))

            sections = cursor.fetchall()

            for section in sections:
                sections_listbox.insert(tk.END, section[0])

        def on_section_select(event=None):
            """عند اختيار فصل من القائمة"""
            selected_indices = sections_listbox.curselection()
            if not selected_indices:
                return

            selected_course = course_var.get().strip()
            selected_section = sections_listbox.get(selected_indices[0])

            if not selected_course or not selected_section:
                return

            # تحديث عنوان التفاصيل
            section_title_var.set(f"فصل: {selected_section}")

            # حساب عدد الطلاب
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT COUNT(*)
                FROM student_sections
                WHERE course_name=? AND section_name=?
            """, (selected_course, selected_section))

            count = cursor.fetchone()[0]
            students_count_var.set(f"عدد الطلاب في فصل \"{selected_section}\": {count}")

            # تعديل: تحديث عرض عدد الطلاب في العنوان
            total_students_var.set(students_count_var.get())

        def add_new_section():
            """إضافة فصل جديد للدورة المحددة"""
            selected_course = course_var.get().strip()

            if not selected_course:
                messagebox.showwarning("تنبيه", "الرجاء اختيار دورة أولاً")
                return

            section_name = simpledialog.askstring("إضافة فصل", "أدخل اسم الفصل الجديد:")

            if not section_name:
                return

            # التحقق من وجود الفصل
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT COUNT(*)
                FROM course_sections
                WHERE course_name=? AND section_name=?
            """, (selected_course, section_name))

            if cursor.fetchone()[0] > 0:
                messagebox.showwarning("تنبيه", f"الفصل '{section_name}' موجود بالفعل في هذه الدورة")
                return

            # إضافة الفصل الجديد
            try:
                current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with self.conn:
                    self.conn.execute("""
                        INSERT INTO course_sections (course_name, section_name, created_date)
                        VALUES (?, ?, ?)
                    """, (selected_course, section_name, current_date))

                messagebox.showinfo("نجاح", f"تم إضافة الفصل '{section_name}' بنجاح")
                update_sections_list()
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء إضافة الفصل: {str(e)}")

        def rename_section():
            """تغيير اسم الفصل المحدد"""
            selected_indices = sections_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("تنبيه", "الرجاء اختيار فصل أولاً")
                return

            selected_course = course_var.get().strip()
            old_section_name = sections_listbox.get(selected_indices[0])

            new_section_name = simpledialog.askstring("تغيير اسم الفصل", "أدخل الاسم الجديد للفصل:",
                                                      initialvalue=old_section_name)

            if not new_section_name or new_section_name == old_section_name:
                return

            # التحقق من وجود الفصل الجديد
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT COUNT(*)
                FROM course_sections
                WHERE course_name=? AND section_name=?
            """, (selected_course, new_section_name))

            if cursor.fetchone()[0] > 0:
                messagebox.showwarning("تنبيه", f"الفصل '{new_section_name}' موجود بالفعل في هذه الدورة")
                return

            # تحديث اسم الفصل
            try:
                with self.conn:
                    # تحديث في جدول الفصول
                    self.conn.execute("""
                        UPDATE course_sections
                        SET section_name=?
                        WHERE course_name=? AND section_name=?
                    """, (new_section_name, selected_course, old_section_name))

                    # تحديث في جدول الطلاب
                    self.conn.execute("""
                        UPDATE student_sections
                        SET section_name=?
                        WHERE course_name=? AND section_name=?
                    """, (new_section_name, selected_course, old_section_name))

                messagebox.showinfo("نجاح", f"تم تغيير اسم الفصل إلى '{new_section_name}' بنجاح")
                update_sections_list()
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء تغيير اسم الفصل: {str(e)}")

        def manage_section_students():
            """إدارة طلاب الفصل المحدد"""
            selected_indices = sections_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("تنبيه", "الرجاء اختيار فصل أولاً")
                return

            selected_course = course_var.get().strip()
            selected_section = sections_listbox.get(selected_indices[0])

            # فتح نافذة إدارة طلاب الفصل
            self.open_section_students_window(selected_course, selected_section)

        def export_section_attendance_sheet():
            """تصدير كشف حضور للفصل المحدد"""
            selected_indices = sections_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("تنبيه", "الرجاء اختيار فصل أولاً")
                return

            selected_course = course_var.get().strip()
            selected_section = sections_listbox.get(selected_indices[0])

            # تنفيذ وظيفة تصدير كشف الحضور للفصل
            self.export_section_to_word(selected_course, selected_section)

        def export_section_diligence():
            """تصدير كشف المواظبة والسلوك للفصل المحدد"""
            selected_indices = sections_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("تنبيه", "الرجاء اختيار فصل أولاً")
                return

            selected_course = course_var.get().strip()
            selected_section = sections_listbox.get(selected_indices[0])

            # تنفيذ وظيفة تصدير كشف المواظبة والسلوك للفصل
            self.export_section_diligence_behavior(selected_course, selected_section)

        # وظيفة حذف الفصل مع ترحيل الطلاب
        def delete_section_with_transfer():
            """حذف الفصل مع ترحيل الطلاب إلى فصل آخر"""
            selected_indices = sections_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("تنبيه", "الرجاء اختيار فصل للحذف")
                return

            selected_course = course_var.get().strip()
            selected_section = sections_listbox.get(selected_indices[0])

            # الحصول على عدد الطلاب في الفصل المحدد
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT COUNT(*)
                FROM student_sections
                WHERE course_name=? AND section_name=?
            """, (selected_course, selected_section))

            students_count = cursor.fetchone()[0]

            # التحقق من وجود فصول أخرى
            cursor.execute("""
                SELECT section_name
                FROM course_sections
                WHERE course_name=? AND section_name!=?
            """, (selected_course, selected_section))

            other_sections = [row[0] for row in cursor.fetchall()]

            if not other_sections:
                messagebox.showwarning("تنبيه",
                                       f"لا يوجد فصول أخرى في الدورة '{selected_course}' لترحيل الطلاب إليها.\n\nيجب أن يتوفر فصل واحد على الأقل لنقل الطلاب إليه.")
                return

            # إذا كان هناك طلاب في الفصل، اطلب تحديد فصل للترحيل
            if students_count > 0:
                # عرض نافذة لاختيار الفصل المراد الترحيل إليه
                transfer_window = tk.Toplevel(multi_window)
                transfer_window.title("ترحيل الطلاب")
                transfer_window.geometry("400x300")
                transfer_window.configure(bg=self.colors["light"])
                transfer_window.transient(multi_window)
                transfer_window.grab_set()

                # توسيط النافذة
                x = (transfer_window.winfo_screenwidth() - 400) // 2
                y = (transfer_window.winfo_screenheight() - 300) // 2
                transfer_window.geometry(f"400x300+{x}+{y}")

                tk.Label(
                    transfer_window,
                    text=f"ترحيل طلاب الفصل: {selected_section}",
                    font=self.fonts["title"],
                    bg=self.colors["primary"],
                    fg="white",
                    padx=10, pady=10
                ).pack(fill=tk.X)

                tk.Label(
                    transfer_window,
                    text=f"يوجد {students_count} طالب في هذا الفصل.\nاختر الفصل المراد ترحيل الطلاب إليه:",
                    font=self.fonts["text"],
                    bg=self.colors["light"],
                    pady=10
                ).pack()

                # قائمة الفصول المتاحة للترحيل
                target_var = tk.StringVar()
                target_listbox = tk.Listbox(
                    transfer_window,
                    font=self.fonts["text"],
                    selectbackground=self.colors["primary"],
                    selectforeground="white",
                    height=8
                )
                target_listbox.pack(fill=tk.X, padx=20, pady=10)

                # إضافة أسماء الفصول إلى القائمة
                for section in other_sections:
                    target_listbox.insert(tk.END, section)

                # إذا كان هناك فصل واحد فقط، حدده تلقائيًا
                if len(other_sections) == 1:
                    target_listbox.select_set(0)

                button_frame = tk.Frame(transfer_window, bg=self.colors["light"], pady=10)
                button_frame.pack(fill=tk.X, padx=20)

                def execute_transfer():
                    """تنفيذ عملية الترحيل وحذف الفصل"""
                    selected_indices = target_listbox.curselection()
                    if not selected_indices:
                        messagebox.showwarning("تنبيه", "الرجاء اختيار فصل للترحيل إليه")
                        return

                    target_section = target_listbox.get(selected_indices[0])

                    try:
                        with self.conn:
                            # ترحيل الطلاب إلى الفصل المحدد
                            self.conn.execute("""
                                UPDATE student_sections
                                SET section_name=?
                                WHERE course_name=? AND section_name=?
                            """, (target_section, selected_course, selected_section))

                            # حذف الفصل
                            self.conn.execute("""
                                DELETE FROM course_sections
                                WHERE course_name=? AND section_name=?
                            """, (selected_course, selected_section))

                        messagebox.showinfo("نجاح",
                                            f"تم ترحيل {students_count} طالب من الفصل '{selected_section}' إلى الفصل '{target_section}' وحذف الفصل بنجاح")
                        transfer_window.destroy()
                        update_sections_list()

                        # تحديث الإحصائيات بعد عملية الترحيل
                        self.update_statistics()
                        self.update_students_tree()
                        self.update_attendance_display()

                    except Exception as e:
                        messagebox.showerror("خطأ", f"حدث خطأ أثناء الترحيل: {str(e)}")

                transfer_btn = tk.Button(
                    button_frame,
                    text="ترحيل وحذف",
                    font=self.fonts["text_bold"],
                    bg=self.colors["warning"],
                    fg="white",
                    padx=15, pady=5,
                    bd=0, relief=tk.FLAT,
                    cursor="hand2",
                    command=execute_transfer
                )
                transfer_btn.pack(side=tk.LEFT)

                cancel_btn = tk.Button(
                    button_frame,
                    text="إلغاء",
                    font=self.fonts["text_bold"],
                    bg=self.colors["danger"],
                    fg="white",
                    padx=15, pady=5,
                    bd=0, relief=tk.FLAT,
                    cursor="hand2",
                    command=transfer_window.destroy
                )
                cancel_btn.pack(side=tk.RIGHT)

            else:
                # إذا لم يكن هناك طلاب، يمكن حذف الفصل مباشرة
                if messagebox.askyesno("تأكيد", f"هل أنت متأكد من حذف الفصل '{selected_section}'؟"):
                    try:
                        with self.conn:
                            self.conn.execute("""
                                DELETE FROM course_sections
                                WHERE course_name=? AND section_name=?
                            """, (selected_course, selected_section))

                        messagebox.showinfo("نجاح", f"تم حذف الفصل '{selected_section}' بنجاح")
                        update_sections_list()

                        # تحديث الإحصائيات بعد حذف الفصل
                        self.update_statistics()
                        self.update_students_tree()
                        self.update_attendance_display()

                    except Exception as e:
                        messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف الفصل: {str(e)}")

        # وظيفة حذف الدورة كاملة (للمشرفين فقط)
        def delete_entire_course():
            """حذف الدورة كاملة مع جميع الفصول والطلاب"""
            if not self.current_user["permissions"]["is_admin"]:
                messagebox.showwarning("تنبيه", "هذه الوظيفة متاحة للمشرفين فقط")
                return

            selected_course = course_var.get().strip()
            if not selected_course:
                messagebox.showwarning("تنبيه", "الرجاء اختيار دورة للحذف")
                return

            # التأكيد قبل الحذف
            confirmation = messagebox.askquestion(
                "تحذير - حذف دورة كاملة",
                f"تحذير! أنت على وشك حذف الدورة '{selected_course}' بالكامل.\n\n"
                "سيؤدي هذا إلى:\n"
                "• حذف جميع الفصول في الدورة\n"
                "• حذف جميع الطلاب المرتبطين بالدورة\n"
                "• حذف جميع سجلات الحضور المرتبطة بالدورة\n\n"
                "هذا الإجراء لا يمكن التراجع عنه.\n\n"
                "هل أنت متأكد من رغبتك في حذف الدورة بالكامل؟",
                icon="warning",
                type="yesnocancel"
            )

            if confirmation != "yes":
                return

            # طلب كلمة مرور المشرف للتأكيد
            admin_password = simpledialog.askstring(
                "تأكيد حذف الدورة",
                "أدخل كلمة مرور المشرف للتأكيد:",
                show="*"
            )

            if not admin_password:
                return

            # التحقق من كلمة المرور
            hashed_password = hashlib.sha256(admin_password.encode()).hexdigest()
            cursor = self.conn.cursor()
            cursor.execute("SELECT password FROM users WHERE username=?", ("admin",))
            result = cursor.fetchone()

            if not result or result[0] != hashed_password:
                messagebox.showwarning("تنبيه", "كلمة المرور غير صحيحة")
                return

            # بدء عملية الحذف
            try:
                # إظهار نافذة تقدم العملية
                progress_window = tk.Toplevel(multi_window)
                progress_window.title("جاري حذف الدورة")
                progress_window.geometry("400x150")
                progress_window.configure(bg=self.colors["light"])
                progress_window.transient(multi_window)
                progress_window.grab_set()

                # توسيط النافذة
                x = (progress_window.winfo_screenwidth() - 400) // 2
                y = (progress_window.winfo_screenheight() - 150) // 2
                progress_window.geometry(f"400x150+{x}+{y}")

                tk.Label(
                    progress_window,
                    text=f"جاري حذف الدورة '{selected_course}'...",
                    font=self.fonts["text_bold"],
                    bg=self.colors["light"],
                    pady=10
                ).pack()

                progress_var = tk.DoubleVar()
                progress_bar = ttk.Progressbar(
                    progress_window,
                    variable=progress_var,
                    maximum=100,
                    length=350
                )
                progress_bar.pack(pady=10)

                status_label = tk.Label(
                    progress_window,
                    text="جاري تحضير العملية...",
                    font=self.fonts["text"],
                    bg=self.colors["light"]
                )
                status_label.pack(pady=5)

                progress_window.update()

                # الحصول على جميع أرقام هويات الطلاب في الدورة
                cursor.execute("""
                    SELECT national_id 
                    FROM trainees 
                    WHERE course=?
                """, (selected_course,))
                student_ids = [row[0] for row in cursor.fetchall()]

                total_steps = 3
                current_step = 0

                # 1. حذف سجلات الحضور
                progress_var.set((current_step / total_steps) * 100)
                status_label.config(text="جاري حذف سجلات الحضور...")
                progress_window.update()

                with self.conn:
                    for student_id in student_ids:
                        self.conn.execute("""
                            DELETE FROM attendance 
                            WHERE national_id=?
                        """, (student_id,))

                current_step += 1
                progress_var.set((current_step / total_steps) * 100)
                status_label.config(text="جاري حذف سجلات الفصول...")
                progress_window.update()

                # 2. حذف سجلات الفصول
                with self.conn:
                    self.conn.execute("""
                        DELETE FROM student_sections 
                        WHERE course_name=?
                    """, (selected_course,))

                    self.conn.execute("""
                        DELETE FROM course_sections 
                        WHERE course_name=?
                    """, (selected_course,))

                current_step += 1
                progress_var.set((current_step / total_steps) * 100)
                status_label.config(text="جاري حذف بيانات الطلاب...")
                progress_window.update()

                # 3. حذف الطلاب
                with self.conn:
                    self.conn.execute("""
                        DELETE FROM trainees 
                        WHERE course=?
                    """, (selected_course,))

                current_step += 1
                progress_var.set(100)
                status_label.config(text="تم حذف الدورة بنجاح!")
                progress_window.update()

                # تحديث الإحصائيات بعد الحذف
                self.update_statistics()
                self.update_students_tree()
                self.update_attendance_display()

                # إغلاق نافذة التقدم بعد ثلاث ثوان
                progress_window.after(3000, progress_window.destroy)

                messagebox.showinfo("نجاح", f"تم حذف الدورة '{selected_course}' بنجاح مع جميع البيانات المرتبطة بها")

                # تحديث القائمة
                cursor.execute("SELECT DISTINCT course_name FROM course_sections")
                updated_courses = [row[0] for row in cursor.fetchall()]
                course_dropdown['values'] = updated_courses

                # مسح القيمة الحالية إذا تم حذفها
                if selected_course not in updated_courses:
                    course_var.set("")

                # تحديث قائمة الفصول
                sections_listbox.delete(0, tk.END)
                section_title_var.set("اختر فصلاً لعرض تفاصيله")
                students_count_var.set("")

            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف الدورة: {str(e)}")
                try:
                    progress_window.destroy()
                except:
                    pass

        # ربط وظيفة اختيار الفصل
        sections_listbox.bind("<<ListboxSelect>>", on_section_select)

        # ربط وظيفة تحديث القائمة بتغيير الدورة
        course_dropdown.bind("<<ComboboxSelected>>", lambda e: update_sections_list())

    def open_section_students_window(self, course_name, section_name):
        """فتح نافذة إدارة طلاب الفصل"""
        students_window = tk.Toplevel(self.root)
        students_window.title(f"إدارة طلاب فصل {section_name} - {course_name}")
        students_window.geometry("900x600")
        students_window.configure(bg=self.colors["light"])
        students_window.grab_set()
        students_window.resizable(True, True)

        # توسيط النافذة
        x = (students_window.winfo_screenwidth() - 900) // 2
        y = (students_window.winfo_screenheight() - 600) // 2
        students_window.geometry(f"900x600+{x}+{y}")

        # عنوان النافذة
        tk.Label(
            students_window,
            text=f"إدارة طلاب فصل: {section_name} - دورة: {course_name}",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        # إطار البحث
        search_frame = tk.Frame(students_window, bg=self.colors["light"], padx=10, pady=10)
        search_frame.pack(fill=tk.X)

        tk.Label(
            search_frame,
            text="البحث عن طالب:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(side=tk.RIGHT, padx=5)

        search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_frame,
            textvariable=search_var,
            font=self.fonts["text"],
            width=25
        )
        search_entry.pack(side=tk.RIGHT, padx=5)

        search_btn = tk.Button(
            search_frame,
            text="بحث",
            font=self.fonts["text_bold"],
            bg=self.colors["secondary"],
            fg="white",
            padx=10, pady=2,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: search_students()
        )
        search_btn.pack(side=tk.RIGHT, padx=5)

        # إطار القوائم المزدوجة
        lists_frame = tk.Frame(students_window, bg=self.colors["light"])
        lists_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # قائمة الطلاب في الفصل
        section_frame = tk.LabelFrame(
            lists_frame,
            text=f"طلاب فصل {section_name}",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=5, pady=5
        )
        section_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        section_scroll = tk.Scrollbar(section_frame)
        section_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        section_students = tk.Listbox(
            section_frame,
            font=self.fonts["text"],
            selectbackground=self.colors["primary"],
            selectforeground="white",
            yscrollcommand=section_scroll.set
        )
        section_students.pack(fill=tk.BOTH, expand=True, pady=5)
        section_scroll.config(command=section_students.yview)

        # القائمة الوسطى للأزرار
        middle_frame = tk.Frame(lists_frame, bg=self.colors["light"], width=100)
        middle_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)

        move_to_other_btn = tk.Button(
            middle_frame,
            text=">>",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=5, pady=2,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: move_to_other_section()
        )
        move_to_other_btn.pack(pady=5)

        move_to_current_btn = tk.Button(
            middle_frame,
            text="<<",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=5, pady=2,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: move_to_current_section()
        )
        move_to_current_btn.pack(pady=5)

        # قائمة الطلاب في الدورة بدون فصل أو في فصول أخرى
        other_frame = tk.LabelFrame(
            lists_frame,
            text="طلاب الدورة الآخرين",
            font=self.fonts["text_bold"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=5, pady=5
        )
        other_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        other_scroll = tk.Scrollbar(other_frame)
        other_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        other_students = tk.Listbox(
            other_frame,
            font=self.fonts["text"],
            selectbackground=self.colors["warning"],
            selectforeground="white",
            yscrollcommand=other_scroll.set
        )
        other_students.pack(fill=tk.BOTH, expand=True, pady=5)
        other_scroll.config(command=other_students.yview)

        # إطار المعلومات
        info_frame = tk.Frame(students_window, bg=self.colors["light"], padx=10, pady=5)
        info_frame.pack(fill=tk.X)

        section_count_var = tk.StringVar(value="عدد طلاب الفصل: 0")
        other_count_var = tk.StringVar(value="عدد الطلاب الآخرين: 0")

        section_count_label = tk.Label(
            info_frame,
            textvariable=section_count_var,
            font=self.fonts["text"],
            bg=self.colors["light"]
        )
        section_count_label.pack(side=tk.RIGHT, padx=10)

        other_count_label = tk.Label(
            info_frame,
            textvariable=other_count_var,
            font=self.fonts["text"],
            bg=self.colors["light"]
        )
        other_count_label.pack(side=tk.LEFT, padx=10)

        # إطار الأزرار السفلي
        button_frame = tk.Frame(students_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=10)

        save_btn = tk.Button(
            button_frame,
            text="حفظ التغييرات",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=lambda: save_changes()
        )
        save_btn.pack(side=tk.LEFT, padx=5)

        close_btn = tk.Button(
            button_frame,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=students_window.destroy
        )
        close_btn.pack(side=tk.RIGHT, padx=5)

        # حفظ التغييرات المؤقتة
        # القواميس تخزن: {الهوية: الاسم}
        current_section_students = {}  # الطلاب في الفصل الحالي
        other_section_students = {}  # الطلاب الآخرين
        modified = False  # هل تم تعديل البيانات

        # الوظائف المساعدة
        def load_students():
            """تحميل بيانات الطلاب"""
            nonlocal current_section_students, other_section_students

            # مسح القوائم
            section_students.delete(0, tk.END)
            other_students.delete(0, tk.END)
            current_section_students.clear()
            other_section_students.clear()

            cursor = self.conn.cursor()

            # 1. الحصول على طلاب الفصل الحالي
            cursor.execute("""
                SELECT t.national_id, t.name
                FROM trainees t
                JOIN student_sections s ON t.national_id = s.national_id
                WHERE t.course=? AND s.section_name=? AND t.is_excluded=0
                ORDER BY t.name
            """, (course_name, section_name))

            for row in cursor.fetchall():
                student_id, student_name = row
                display_text = f"{student_name} ({student_id})"
                section_students.insert(tk.END, display_text)
                current_section_students[student_id] = student_name

            # 2. الحصول على باقي طلاب الدورة (غير مسجلين في فصول أو في فصول أخرى)
            cursor.execute("""
                SELECT t.national_id, t.name, 
                       (SELECT section_name FROM student_sections 
                        WHERE national_id=t.national_id AND course_name=t.course) as section
                FROM trainees t
                WHERE t.course=? AND t.is_excluded=0
                ORDER BY t.name
            """, (course_name,))

            for row in cursor.fetchall():
                student_id, student_name, student_section = row

                # تخطي الطلاب في الفصل الحالي
                if student_section == section_name:
                    continue

                # إضافة الطلاب الآخرين
                display_text = f"{student_name} ({student_id})"
                if student_section:
                    display_text += f" - فصل: {student_section}"
                else:
                    display_text += " - بدون فصل"

                other_students.insert(tk.END, display_text)
                other_section_students[student_id] = student_name

            # تحديث العدادات
            section_count_var.set(f"عدد طلاب الفصل: {len(current_section_students)}")
            other_count_var.set(f"عدد الطلاب الآخرين: {len(other_section_students)}")

        def search_students():
            """البحث عن طلاب"""
            search_text = search_var.get().strip()
            if not search_text:
                load_students()
                return

            # مسح القوائم
            section_students.delete(0, tk.END)
            other_students.delete(0, tk.END)

            # البحث في قائمة طلاب الفصل الحالي
            for student_id, student_name in current_section_students.items():
                if (search_text.lower() in student_name.lower() or
                        search_text in student_id):
                    display_text = f"{student_name} ({student_id})"
                    section_students.insert(tk.END, display_text)

            # البحث في قائمة الطلاب الآخرين
            for student_id, student_name in other_section_students.items():
                if (search_text.lower() in student_name.lower() or
                        search_text in student_id):
                    display_text = f"{student_name} ({student_id})"

                    # التحقق من وجود معلومات الفصل
                    cursor = self.conn.cursor()
                    cursor.execute("""
                        SELECT section_name FROM student_sections
                        WHERE national_id=? AND course_name=?
                    """, (student_id, course_name))

                    result = cursor.fetchone()

                    if result and result[0]:
                        display_text += f" - فصل: {result[0]}"
                    else:
                        display_text += " - بدون فصل"

                    other_students.insert(tk.END, display_text)

        def move_to_other_section():
            """نقل الطلاب المحددين من الفصل الحالي إلى قائمة الطلاب الآخرين"""
            nonlocal modified

            selected_indices = section_students.curselection()
            if not selected_indices:
                return

            for index in reversed(selected_indices):
                student_text = section_students.get(index)
                student_id = extract_id_from_text(student_text)

                if student_id in current_section_students:
                    student_name = current_section_students[student_id]

                    # نقل الطالب إلى القائمة الأخرى
                    other_students.insert(tk.END, f"{student_name} ({student_id}) - بدون فصل")
                    other_section_students[student_id] = student_name

                    # حذف الطالب من القائمة الحالية
                    del current_section_students[student_id]
                    section_students.delete(index)

                    modified = True

            # تحديث العدادات
            section_count_var.set(f"عدد طلاب الفصل: {len(current_section_students)}")
            other_count_var.set(f"عدد الطلاب الآخرين: {len(other_section_students)}")

        def move_to_current_section():
            """نقل الطلاب المحددين من قائمة الطلاب الآخرين إلى الفصل الحالي"""
            nonlocal modified

            selected_indices = other_students.curselection()
            if not selected_indices:
                return

            for index in reversed(selected_indices):
                student_text = other_students.get(index)
                student_id = extract_id_from_text(student_text)

                if student_id in other_section_students:
                    student_name = other_section_students[student_id]

                    # نقل الطالب إلى الفصل الحالي
                    section_students.insert(tk.END, f"{student_name} ({student_id})")
                    current_section_students[student_id] = student_name

                    # حذف الطالب من القائمة الأخرى
                    del other_section_students[student_id]
                    other_students.delete(index)

                    modified = True

            # تحديث العدادات
            section_count_var.set(f"عدد طلاب الفصل: {len(current_section_students)}")
            other_count_var.set(f"عدد الطلاب الآخرين: {len(other_section_students)}")

        def extract_id_from_text(text):
            """استخراج رقم الهوية من النص المعروض"""
            # النص بشكل: "اسم الطالب (رقم الهوية) - معلومات إضافية"
            try:
                start = text.find("(") + 1
                end = text.find(")")
                if start > 0 and end > start:
                    return text[start:end]
            except:
                pass
            return ""

        def save_changes():
            """حفظ التغييرات في قاعدة البيانات"""
            nonlocal modified

            if not modified:
                messagebox.showinfo("معلومات", "لم يتم إجراء أي تغييرات")
                return

            # تحديث بيانات الطلاب في قاعدة البيانات
            current_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            try:
                cursor = self.conn.cursor()
                with self.conn:
                    # 1. حذف كل الطلاب من الفصل الحالي
                    self.conn.execute("""
                        DELETE FROM student_sections
                        WHERE course_name=? AND section_name=?
                    """, (course_name, section_name))

                    # 2. إضافة الطلاب الحاليين في الفصل
                    for student_id in current_section_students:
                        self.conn.execute("""
                            INSERT OR REPLACE INTO student_sections
                            (national_id, course_name, section_name, assigned_date)
                            VALUES (?, ?, ?, ?)
                        """, (student_id, course_name, section_name, current_date))

                messagebox.showinfo("نجاح", "تم حفظ التغييرات بنجاح")
                modified = False
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ التغييرات: {str(e)}")

        # تحميل بيانات الطلاب عند فتح النافذة
        load_students()


    def backup_database(self):
        """إنشاء نسخة احتياطية من قاعدة البيانات"""
        if not self.current_user["permissions"]["is_admin"]:
            messagebox.showwarning("تنبيه", "هذه الوظيفة متاحة للمشرفين فقط")
            return

        import shutil
        import datetime
        import os

        # إنشاء مجلد للنسخ الاحتياطي إذا لم يكن موجوداً
        backup_dir = "backup"
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)

        # إنشاء اسم ملف النسخة الاحتياطية مع التاريخ والوقت
        current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(backup_dir, f"attendance_backup_{current_time}.db")

        # إغلاق الاتصال بقاعدة البيانات مؤقتاً
        self.conn.commit()

        try:
            # نسخ ملف قاعدة البيانات
            shutil.copy2("attendance.db", backup_file)
            messagebox.showinfo("نجاح", f"تم إنشاء نسخة احتياطية في: {backup_file}")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء إنشاء النسخة الاحتياطية: {str(e)}")

    def restore_database(self):
        """استرداد نسخة احتياطية من قاعدة البيانات"""
        if not self.current_user["permissions"]["is_admin"]:
            messagebox.showwarning("تنبيه", "هذه الوظيفة متاحة للمشرفين فقط")
            return

        import os
        import shutil
        import sys
        import datetime

        # التحقق من وجود مجلد النسخ الاحتياطي
        backup_dir = "backup"
        if not os.path.exists(backup_dir):
            messagebox.showwarning("تنبيه", "لا يوجد مجلد للنسخ الاحتياطية")
            return

        # الحصول على قائمة ملفات النسخ الاحتياطية
        backup_files = [f for f in os.listdir(backup_dir) if f.startswith("attendance_backup_") and f.endswith(".db")]

        if not backup_files:
            messagebox.showwarning("تنبيه", "لا توجد نسخ احتياطية متاحة")
            return

        # إنشاء نافذة لاختيار النسخة الاحتياطية
        restore_window = tk.Toplevel(self.root)
        restore_window.title("استرداد نسخة احتياطية")
        restore_window.geometry("600x400")
        restore_window.configure(bg=self.colors["light"])
        restore_window.transient(self.root)
        restore_window.grab_set()

        # توسيط النافذة
        x = (restore_window.winfo_screenwidth() - 600) // 2
        y = (restore_window.winfo_screenheight() - 400) // 2
        restore_window.geometry(f"600x400+{x}+{y}")

        # عنوان النافذة
        tk.Label(
            restore_window,
            text="استرداد نسخة احتياطية من قاعدة البيانات",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10, pady=10
        ).pack(fill=tk.X)

        # إطار القائمة
        list_frame = tk.Frame(restore_window, bg=self.colors["light"])
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(
            list_frame,
            text="اختر النسخة الاحتياطية التي تريد استردادها:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(anchor=tk.W, pady=(0, 10))

        # إنشاء قائمة النسخ الاحتياطية
        backup_listbox = tk.Listbox(
            list_frame,
            font=self.fonts["text"],
            selectbackground=self.colors["primary"],
            selectforeground="white"
        )
        backup_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # تعبئة القائمة بالملفات مرتبة من الأحدث إلى الأقدم
        backup_files.sort(reverse=True)
        for file in backup_files:
            # استخراج التاريخ والوقت من اسم الملف
            date_str = file.replace("attendance_backup_", "").replace(".db", "")
            try:
                # محاولة تنسيق التاريخ بشكل مقروء
                date_year = date_str[:4]
                date_month = date_str[4:6]
                date_day = date_str[6:8]
                time_hour = date_str[9:11]
                time_min = date_str[11:13]
                time_sec = date_str[13:15]
                display_date = f"{date_year}-{date_month}-{date_day} {time_hour}:{time_min}:{time_sec}"
            except:
                display_date = date_str

            backup_listbox.insert(tk.END, f"{display_date} - {file}")

        # إطار الأزرار
        button_frame = tk.Frame(restore_window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=10)

        # تخزين المتغيرات المطلوبة للدالة الداخلية
        def do_restore():
            """تنفيذ عملية استرداد النسخة الاحتياطية المحددة"""
            selection = backup_listbox.curselection()
            if not selection:
                messagebox.showwarning("تنبيه", "الرجاء اختيار نسخة احتياطية")
                return

            selected_item = backup_listbox.get(selection[0])
            backup_file = selected_item.split(" - ")[1]
            backup_path = os.path.join(backup_dir, backup_file)

            # التأكيد على استرداد النسخة الاحتياطية
            if not messagebox.askokcancel(
                    "تأكيد الاسترداد",
                    "سيتم استبدال قاعدة البيانات الحالية بالنسخة الاحتياطية المحددة.\n\n"
                    "هذه العملية ستؤدي إلى فقدان أي تغييرات تمت منذ إنشاء هذه النسخة.\n\n"
                    "هل أنت متأكد من المتابعة؟",
                    icon="warning"
            ):
                return

            # إغلاق الاتصال بقاعدة البيانات
            self.conn.close()

            try:
                # إنشاء نسخة احتياطية إضافية من قاعدة البيانات الحالية (للأمان)
                current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                safety_backup = os.path.join(backup_dir, f"pre_restore_{current_time}.db")

                # نسخ ملف قاعدة البيانات الحالية كإجراء وقائي
                try:
                    shutil.copy2("attendance.db", safety_backup)
                except:
                    pass  # تجاهل الأخطاء في النسخ الوقائي

                # استبدال قاعدة البيانات الحالية بالنسخة الاحتياطية
                shutil.copy2(backup_path, "attendance.db")

                messagebox.showinfo(
                    "نجاح",
                    f"تم استرداد النسخة الاحتياطية بنجاح.\n\n"
                    f"ستتم إعادة تشغيل النظام الآن لتطبيق التغييرات."
                )

                # إغلاق نافذة الاسترداد
                restore_window.destroy()

                # إعادة تشغيل البرنامج
                self.root.destroy()
                python = sys.executable
                os.execl(python, python, *sys.argv)

            except Exception as e:
                messagebox.showerror(
                    "خطأ",
                    f"حدث خطأ أثناء استرداد النسخة الاحتياطية:\n{str(e)}\n\n"
                    "الرجاء إعادة تشغيل البرنامج يدوياً."
                )
                # محاولة إعادة فتح الاتصال بقاعدة البيانات في حال وجود خطأ
                try:
                    self.conn = sqlite3.connect("attendance.db")
                except:
                    pass

        restore_btn = tk.Button(
            button_frame,
            text="استرداد النسخة المحددة",
            font=self.fonts["text_bold"],
            bg=self.colors["warning"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=do_restore
        )
        restore_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = tk.Button(
            button_frame,
            text="إلغاء",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15, pady=5,
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            command=restore_window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)

    def export_student_absence_reports(self, student_info, attendance_records):
        """
        تصدير محاضر غيابات الطالب (محضر منفصل لكل يوم غياب)

        Args:
            student_info: معلومات الطالب (national_id, name, rank, course, ...)
            attendance_records: سجلات الحضور والغياب للطالب
        """
        try:
            # التأكد من وجود مكتبة python-docx
            if 'Document' not in globals():
                messagebox.showerror("خطأ",
                                     "لم يتم العثور على مكتبة python-docx. قم بتثبيتها باستخدام: pip install python-docx")
                return

            # استخراج المعلومات من البيانات
            nid = student_info[0]
            name = student_info[1]
            rank = student_info[2]
            course = student_info[3]

            # البحث عن سجلات الغياب فقط وترتيبها من الأقدم للأحدث
            absence_records = []
            for record in attendance_records:
                status = record[7]  # Status column
                if status == "غائب":
                    absence_records.append(record)

            # ترتيب سجلات الغياب من الأقدم للأحدث
            absence_records = sorted(absence_records, key=lambda x: x[6])

            if not absence_records:
                messagebox.showinfo("معلومات", "لا توجد سجلات غياب لهذا الطالب")
                return

            # إنشاء مستند جديد
            doc = Document()

            # إعداد المستند للغة العربية (RTL)
            section = doc.sections[0]
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)
            section.top_margin = Inches(1.0)
            section.bottom_margin = Inches(1.0)

            # لكل سجل غياب، إنشاء صفحة جديدة
            for i, record in enumerate(absence_records):
                if i > 0:
                    doc.add_page_break()

                # تاريخ الغياب
                absence_date = record[6]  # Date column

                # إضافة عنوان المستند
                title = doc.add_heading('محضر غياب', level=0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in title.runs:
                    run.font.size = Pt(18)
                    run.font.bold = True
                    run.font.rtl = True

                # إضافة تاريخ الغياب تحت العنوان
                date_paragraph = doc.add_paragraph()
                date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                date_run = date_paragraph.add_run(f"تاريخ الغياب: {absence_date}")
                date_run.font.size = Pt(14)
                date_run.font.bold = True
                date_run.font.rtl = True

                # إضافة خط أفقي
                doc.add_paragraph().paragraph_format.border_bottom = True

                # إضافة جدول معلومات الطالب
                doc.add_paragraph()  # فراغ قبل الجدول
                student_table = doc.add_table(rows=1, cols=4)
                student_table.style = 'Table Grid'

                # إضافة عناوين الجدول
                header_cells = student_table.rows[0].cells

                # نظراً لأن اللغة العربية RTL، نضيف العناوين بشكل معكوس
                header_cells[3].text = "الاسم"
                header_cells[2].text = "الرتبة"
                header_cells[1].text = "رقم الهوية"
                header_cells[0].text = "الدورة"

                # تنسيق عناوين الجدول
                for cell in header_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True
                            run.font.rtl = True
                            run.font.size = Pt(12)

                    # إضافة تظليل للرأس
                    try:
                        shading_elm = parse_xml(r'<w:shd {} w:fill="DDDDDD"/>'.format(nsdecls('w')))
                        cell._element.get_or_add_tcPr().append(shading_elm)
                    except:
                        pass

                # إضافة بيانات الطالب
                data_row = student_table.add_row().cells

                # نضيف البيانات بشكل معكوس بسبب RTL
                data_row[3].text = name
                data_row[2].text = rank
                data_row[1].text = nid
                data_row[0].text = course

                # تنسيق بيانات الطالب
                for cell in data_row:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.rtl = True
                            run.font.size = Pt(12)

                # ضبط عرض الجدول
                student_table.autofit = False
                try:
                    col_widths = [2.5, 1.5, 1.5, 2.0]  # العرض بالبوصة (الدورة، الهوية، الرتبة، الاسم)
                    for i, width in enumerate(col_widths):
                        student_table.columns[i].width = Inches(width)
                except:
                    pass

                # إضافة نص المحضر
                doc.add_paragraph()  # فراغ بعد الجدول

                absence_paragraph = doc.add_paragraph()
                absence_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                absence_paragraph.paragraph_format.space_after = Pt(12)
                absence_text = f"أنه في يوم {self.get_arabic_day_name(absence_date)} الموافق {absence_date} قد تغيب المتدرب المذكور أعلاه عن دورة {course} وبناءً عليه أعد هذا المحضر للاطلاع."
                absence_run = absence_paragraph.add_run(absence_text)
                absence_run.font.rtl = True
                absence_run.font.size = Pt(12)

                # إضافة مكان للتوقيعات
                doc.add_paragraph()
                doc.add_paragraph()

                signature_table = doc.add_table(rows=1, cols=3)
                signature_table.style = 'Table Grid'

                sig_cells = signature_table.rows[0].cells
                sig_cells[2].text = "مسؤول الحضور: _________________"
                sig_cells[1].text = "رئيس القسم: __________________"
                sig_cells[0].text = "مدير التدريب: ________________"

                for cell in sig_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in cell.paragraphs[0].runs:
                        run.font.rtl = True
                        run.font.size = Pt(11)

                # إضافة تاريخ الطباعة
                doc.add_paragraph()

                now = datetime.datetime.now()
                date_print_str = now.strftime("%Y-%m-%d")

                print_date = doc.add_paragraph()
                print_date.alignment = WD_ALIGN_PARAGRAPH.LEFT
                print_date_run = print_date.add_run(f"تاريخ الطباعة: {date_print_str}")
                print_date_run.font.size = Pt(9)
                print_date_run.font.rtl = True

            # حفظ المستند
            export_file = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile=f"محاضر_غياب_{name}_{nid}.docx"
            )

            if export_file:
                doc.save(export_file)
                messagebox.showinfo("نجاح",
                                    f"تم تصدير محاضر الغياب ({len(absence_records)} محضر) بنجاح إلى:\n{export_file}")

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تصدير محاضر الغياب: {str(e)}")

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



# =============================================================================
#                      نقطة التشغيل الرئيسية
# =============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = LoginSystem(root)
    root.mainloop()
