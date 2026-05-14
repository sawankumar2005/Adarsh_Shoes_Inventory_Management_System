import os
from datetime import datetime
import pandas as pd
import customtkinter
from tkinter import filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import qrcode
from PIL import Image, ImageTk
import threading

# ---- Config ----
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

EXCEL_FILE_STOCK = "stock.xlsx"
EXCEL_FILE_SALES = "sales_history.xlsx"
LOW_STOCK_THRESHOLD = 5

# ---- Helper Functions ----

def run_in_thread(func, callback, *args, **kwargs):
    """Runs a function in a separate thread and calls a callback with the result."""
    def wrapper():
        try:
            result = func(*args, **kwargs)
            # Use after(0, ...) to schedule the callback on the main thread
            app.after(0, lambda: callback(result))
        except Exception as e:
            app.after(0, lambda: callback(("Error", f"An error occurred: {e}")))

    thread = threading.Thread(target=wrapper)
    thread.daemon = True
    thread.start()

def sg_msgbox(title, message):
    msg_box = customtkinter.CTkToplevel()
    msg_box.title(title)
    msg_box.geometry("420x160")
    msg_box.resizable(False, False)
    customtkinter.CTkLabel(msg_box, text=message, font=("Arial", 12), wraplength=380).pack(pady=(20, 10), padx=12)
    ok_btn = customtkinter.CTkButton(msg_box, text="OK", width=80, command=msg_box.destroy)
    ok_btn.pack(pady=(6, 12))
    msg_box.grab_set()
    msg_box.wait_window()

def sg_confirm(title, message):
    result = {'value': False}

    def on_ok():
        result['value'] = True
        dlg.destroy()

    def on_cancel():
        dlg.destroy()

    dlg = customtkinter.CTkToplevel()
    dlg.title(title)
    dlg.geometry("480x160")
    dlg.resizable(False, False)
    customtkinter.CTkLabel(dlg, text=message, font=("Arial", 12), wraplength=440).pack(pady=(18, 10), padx=12)
    btn_frame = customtkinter.CTkFrame(dlg, fg_color="transparent")
    btn_frame.pack(pady=8)
    customtkinter.CTkButton(btn_frame, text="OK", width=80, command=on_ok).pack(side="left", padx=8)
    customtkinter.CTkButton(btn_frame, text="Cancel", width=80, command=on_cancel).pack(side="left", padx=8)
    dlg.grab_set()
    dlg.wait_window()
    return result['value']

# ---- File read/write helpers ----
def read_stock():
    if os.path.exists(EXCEL_FILE_STOCK):
        try:
            df = pd.read_excel(EXCEL_FILE_STOCK, dtype={'item_id': str, 'unique_id': str}, engine='openpyxl')
            expected = ['item_id', 'name', 'size', 'color', 'article', 'quantity', 'unique_id']
            for c in expected:
                if c not in df.columns:
                    df[c] = "" if c != 'quantity' else 0
            df['quantity'] = df['quantity'].fillna(0).astype(int)
            df['unique_id'] = df['unique_id'].astype(str)
            return df[expected]
        except Exception as e:
            print("ERROR reading stock file:", e)
    return pd.DataFrame(columns=['item_id', 'name', 'size', 'color', 'article', 'quantity', 'unique_id'])

def write_stock(df):
    try:
        df.to_excel(EXCEL_FILE_STOCK, index=False, engine='openpyxl')
    except Exception as e:
        print("ERROR writing stock file:", e)
        sg_msgbox("File Write Error", f"Could not write stock file:\n{e}")

def read_sales():
    if os.path.exists(EXCEL_FILE_SALES):
        try:
            df = pd.read_excel(EXCEL_FILE_SALES, engine='openpyxl')
            if 'qty' not in df.columns:
                df['qty'] = 1
            df['qty'] = df['qty'].fillna(0).astype(int)
            if 'type' not in df.columns:
                df['type'] = 'sale'
            if 'date' in df.columns:
                df['date'] = pd.to_datetime(df['date'], errors='coerce').dt.strftime('%Y-%m-%d')
            else:
                df['date'] = datetime.now().strftime('%Y-%m-%d')
            if 'floor' not in df.columns:
                df['floor'] = 'All Floors'
            for c in ['unique_id', 'date', 'time', 'name', 'size', 'qty', 'type', 'floor']:
                if c not in df.columns:
                    df[c] = ""
            return df[['unique_id', 'date', 'time', 'name', 'size', 'qty', 'type', 'floor']]
        except Exception as e:
            print("ERROR reading sales file:", e)
    return pd.DataFrame(columns=['unique_id', 'date', 'time', 'name', 'size', 'qty', 'type', 'floor'])

def write_sales(df):
    try:
        df.to_excel(EXCEL_FILE_SALES, index=False, engine='openpyxl')
    except Exception as e:
        print("ERROR writing sales file:", e)
        sg_msgbox("File Write Error", f"Could not write sales file:\n{e}")


# ---- Business logic ----
def process_stock_action(item_id, action, item_details=None, qty=1):
    df = read_stock()
    df_sales = read_sales()

    title, msg = "Info", "No action taken."
    stock_before_return = None
    stock_after_return = None

    if action == 'add':
        if item_details is None:
            return "Error", "Missing item details to add stock."

        unique_id = f"{item_details['name'].replace(' ', '-')}-{item_details['article'].replace(' ', '-')}-{item_details['size']}-{item_details['color'].replace(' ', '-')}"
        
        if unique_id in df['unique_id'].values:
            idx = df[df['unique_id'] == unique_id].index[0]
            df.at[idx, 'quantity'] = int(df.at[idx, 'quantity']) + int(item_details['quantity'])
            title, msg = "Stock Updated", f"{item_details['quantity']} added. New qty: {df.at[idx,'quantity']}"
        else:
            new_row = {
                'item_id': item_details['name'],
                'name': item_details['name'],
                'size': item_details['size'],
                'color': item_details['color'],
                'article': item_details['article'],
                'quantity': int(item_details['quantity']),
                'unique_id': unique_id
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            title, msg = "New Item", f"Added {unique_id} with qty {item_details['quantity']}"

    elif action == 'sell':
        item_id = item_id.replace(' ', '-')
        if item_id in df['unique_id'].values:
            df = df.copy() 
            idx = df[df['unique_id'] == item_id].index[0]
            current_qty = int(df.at[idx, 'quantity'])
            if current_qty >= qty:
                df.at[idx, 'quantity'] = current_qty - qty
                item_data = df.loc[idx]
                
                floor = item_details.get('floor', 'All Floors') if item_details else 'All Floors'

                new_sales_entry = {
                    'unique_id': item_id,
                    'date': datetime.now().strftime("%Y-%m-%d"),
                    'time': datetime.now().strftime("%H:%M:%S"),
                    'name': item_data['name'],
                    'size': item_data['size'],
                    'qty': qty,
                    'type': 'sale',
                    'floor': floor
                }
                df_sales = pd.concat([df_sales, pd.DataFrame([new_sales_entry])], ignore_index=True)

                msg = f"Sold {qty} of {item_id}. Remaining: {int(df.at[idx,'quantity'])}"
                title = "Item Sold"

                if int(df.at[idx, 'quantity']) <= LOW_STOCK_THRESHOLD:
                    msg += f"\n\n⚠️ Low Stock Alert! The quantity of {item_data['name']} is now {int(df.at[idx,'quantity'])}. Please re-order soon!"
            else:
                title, msg = "Error", f"Only {current_qty} in stock!"
        else:
            title, msg = "Error", f"{item_id} not found."

    elif action == 'return':
        item_id = item_id.replace(' ', '-')
        if item_id in df['unique_id'].values:
            df = df.copy()
            idx = df[df['unique_id'] == item_id].index[0]
            
            stock_before_return = int(df.at[idx, 'quantity'])
            df.at[idx, 'quantity'] = stock_before_return + int(qty)
            stock_after_return = int(df.at[idx, 'quantity'])

            item_data = df.loc[idx]
            floor = item_details.get('floor', 'All Floors') if item_details else 'All Floors'
            new_return_entry = {
                'unique_id': item_id,
                'date': datetime.now().strftime("%Y-%m-%d"),
                'time': datetime.now().strftime("%H:%M:%S"),
                'name': item_data['name'],
                'size': item_data['size'],
                'qty': int(qty),
                'type': 'return',
                'floor': floor
            }
            df_sales = pd.concat([df_sales, pd.DataFrame([new_return_entry])], ignore_index=True)

            result = ("Item Returned", stock_before_return, stock_after_return)
            write_stock(df)
            write_sales(df_sales)
            return result
        else:
            title, msg = "Error", "Item not found in stock."

    write_stock(df)
    write_sales(df_sales)
    return title, msg


# ---- GUI Frames (Moved to a class for better state management) ----
class DashboardFrame(customtkinter.CTkFrame):
    def __init__(self, master, app_instance):
        super().__init__(master)
        
        self.app_instance = app_instance
        self.grid_rowconfigure(0, weight=1)

        header_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, pady=(14, 0), padx=20, sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1)

        self.date_label = customtkinter.CTkLabel(header_frame, text="", font=("Arial", 14))
        self.date_label.grid(row=0, column=0, sticky="w")

        customtkinter.CTkLabel(self, text="👟 Adarsh Shoes Inventory Dashboard", font=("Arial", 22, "bold")).grid(row=1, column=0, pady=(6, 12), sticky="ew")

        card_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        card_frame.grid(row=2, column=0, pady=6, sticky="ew")
        card_frame.grid_columnconfigure(0, weight=1)
        card_frame.grid_columnconfigure(1, weight=1)
        card_frame.grid_columnconfigure(2, weight=1)
        card_frame.grid_rowconfigure(0, weight=1)

        def create_card(parent, title, value_attr, bg_color, emoji, col):
            card = customtkinter.CTkFrame(parent, corner_radius=12, fg_color=bg_color)
            card.grid(row=0, column=col, padx=12, ipadx=16, ipady=12, sticky="ew")
            customtkinter.CTkLabel(card, text=f"{emoji} {title}", text_color="white").pack()
            label = customtkinter.CTkLabel(card, text="0", font=("Arial", 20, "bold"), text_color="white")
            label.pack()
            setattr(self, value_attr, label)

        create_card(card_frame, "Total Stock", "total_stock_label", "#2a9d8f", "📦", 0)
        create_card(card_frame, "Today's Sales", "total_sales_label", "#e76f51", "💰", 1)
        create_card(card_frame, "Today's Returns", "total_returns_label", "#f4a261", "↩", 2)

        self.button_container = customtkinter.CTkFrame(self, fg_color="transparent")
        self.button_container.grid(row=3, column=0, pady=10, sticky="ew")
        self.button_container.grid_columnconfigure((0, 1, 2), weight=1)

        self.see_stock_button = customtkinter.CTkButton(self.button_container, text="👀 See Current Stock")
        self.see_stock_button.grid(row=0, column=0, padx=10, sticky="ew")
        self.see_stock_button.configure(command=self.show_current_stock_list)
        
        self.show_low_stock_button = customtkinter.CTkButton(self.button_container, text="⚠️ See Low Stock")
        self.show_low_stock_button.grid(row=0, column=1, padx=10, sticky="ew")
        self.show_low_stock_button.configure(command=self.show_low_stock_list)
        
        # This button will now open the graph options
        self.show_chart_button = customtkinter.CTkButton(self.button_container, text="📊 Show Sales Graph", command=self.show_graph_options)
        self.show_chart_button.grid(row=0, column=2, padx=10, sticky="ew")
        
        self.list_container = customtkinter.CTkFrame(self, fg_color="transparent")
        self.list_container.grid(row=4, column=0, pady=10, padx=20, sticky="nsew")
        self.grid_rowconfigure(4, weight=1)
        self.list_container.grid_columnconfigure(0, weight=1)
        
        self.update_data()
    
    def update_data(self):
        self.date_label.configure(text=datetime.now().strftime("%d %B %Y"))

        df_stock = read_stock()
        total_stock = int(df_stock['quantity'].sum()) if not df_stock.empty else 0
        self.total_stock_label.configure(text=str(total_stock))

        df_sales = read_sales()
        today = datetime.now().strftime('%Y-%m-%d')
        if not df_sales.empty and 'date' in df_sales.columns:
            today_sales = df_sales[(df_sales['date'] == today) & (df_sales['type'] == 'sale')]
            today_returns = df_sales[(df_sales['date'] == today) & (df_sales['type'] == 'return')]
            
            self.total_sales_label.configure(text=str(int(today_sales['qty'].sum()) if not today_sales.empty else 0))
            self.total_returns_label.configure(text=str(int(today_returns['qty'].sum()) if not today_returns.empty else 0))
            
        else:
            self.total_sales_label.configure(text="0")
            self.total_returns_label.configure(text="0")
            
        self.clear_list_area()
    
    def clear_list_area(self):
        for widget in self.list_container.winfo_children():
            widget.destroy()
        self.button_container.grid(row=3, column=0, pady=10, sticky="ew")
        
    def show_list(self, df_list, title, is_low_stock=False):
        self.button_container.grid_forget()
        self.clear_list_area()
        
        customtkinter.CTkLabel(self.list_container, text=title, font=("Arial", 14, "bold")).grid(row=0, column=0, pady=5)
        
        list_frame = customtkinter.CTkScrollableFrame(self.list_container, width=900)
        list_frame.grid(row=1, column=0, pady=10, padx=10, sticky="nsew")
        
        if df_list.empty:
            customtkinter.CTkLabel(list_frame, text="No items found.").grid(row=0, column=0, pady=20)
        else:
            headers = ["Name", "Size", "Color", "Qty", "Article", "Unique ID"]
            for i, h in enumerate(headers):
                customtkinter.CTkLabel(list_frame, text=h, font=("Arial", 12, "bold")).grid(row=0, column=i, padx=5, pady=5)
            
            for i, (_, row) in enumerate(df_list.iterrows()):
                data = [row['name'], row['size'], row['color'], int(row['quantity']), row['article'], row['unique_id']]
                for j, val in enumerate(data):
                    label = customtkinter.CTkLabel(list_frame, text=str(val))
                    if is_low_stock:
                        label.configure(text_color="red" if j == 3 else "black")
                    label.grid(row=i+1, column=j, padx=5, pady=3, sticky="w")
        
        customtkinter.CTkButton(self.list_container, text="Hide List", command=self.clear_list_area).grid(row=2, column=0, pady=5)
    
    def show_current_stock_list(self):
        df_stock = read_stock()
        self.show_list(df_stock, "Current Stock Inventory")

    def show_low_stock_list(self):
        df_stock = read_stock()
        low_stock_items = df_stock[df_stock['quantity'] <= LOW_STOCK_THRESHOLD]
        self.show_list(low_stock_items, f"Low Stock Items (<= {LOW_STOCK_THRESHOLD})", is_low_stock=True)

    def show_graph_options(self):
        self.button_container.grid_forget()
        self.clear_list_area()

        options_frame = customtkinter.CTkFrame(self.list_container, fg_color="transparent")
        options_frame.grid(row=0, column=0, pady=10, padx=20)
        options_frame.grid_columnconfigure(0, weight=1)

        customtkinter.CTkLabel(options_frame, text="Select a Sales Graph Timeframe:", font=("Arial", 16, "bold")).pack(pady=10)
        
        # New buttons for different time ranges
        customtkinter.CTkButton(options_frame, text="📈 Today's Sales", command=lambda: self.show_sales_graph('today')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 This Week's Sales", command=lambda: self.show_sales_graph('this_week')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 This Month's Sales", command=lambda: self.show_sales_graph('this_month')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 Last 3 Months", command=lambda: self.show_sales_graph('last_3_months')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 Last 6 Months", command=lambda: self.show_sales_graph('last_6_months')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 Last 9 Months", command=lambda: self.show_sales_graph('last_9_months')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 This Year's Sales", command=lambda: self.show_sales_graph('this_year')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 Last 2 Years", command=lambda: self.show_sales_graph('last_2_years')).pack(pady=5)
        customtkinter.CTkButton(options_frame, text="📈 Last 5 Years", command=lambda: self.show_sales_graph('last_5_years')).pack(pady=5)

        customtkinter.CTkButton(options_frame, text="❌ Cancel", command=self.clear_list_area).pack(pady=20)


    def show_sales_graph(self, time_range):
        self.clear_list_area()
        
        df_sales = read_sales()
        if df_sales.empty:
            customtkinter.CTkLabel(self.list_container, text="No sales data available").grid(row=0, column=0, pady=10)
            customtkinter.CTkButton(self.list_container, text="Back", command=self.show_graph_options).grid(row=1, column=0, pady=6)
            return

        df_sales['date_parsed'] = pd.to_datetime(df_sales['date'], errors='coerce')

        today = datetime.now().date()
        if time_range == 'today':
            start_date = today
            title = "Today's Sales"
        elif time_range == 'this_week':
            start_date = today - pd.Timedelta(days=today.weekday())
            title = "This Week's Sales"
        elif time_range == 'this_month':
            start_date = today.replace(day=1)
            title = "This Month's Sales"
        elif time_range == 'last_3_months':
            start_date = today - pd.Timedelta(days=90)
            title = "Sales from Last 3 Months"
        elif time_range == 'last_6_months':
            start_date = today - pd.Timedelta(days=180)
            title = "Sales from Last 6 Months"
        elif time_range == 'last_9_months':
            start_date = today - pd.Timedelta(days=270)
            title = "Sales from Last 9 Months"
        elif time_range == 'this_year':
            start_date = today.replace(month=1, day=1)
            title = "This Year's Sales"
        elif time_range == 'last_2_years':
            start_date = today - pd.Timedelta(days=365*2)
            title = "Sales from Last 2 Years"
        elif time_range == 'last_5_years':
            start_date = today - pd.Timedelta(days=365*5)
            title = "Sales from Last 5 Years"
        else: # Default to last 7 days if something goes wrong
            start_date = today - pd.Timedelta(days=7)
            title = "Sales from Last 7 Days"
        
        filtered_df = df_sales[df_sales['date_parsed'] >= pd.to_datetime(start_date)]
        
        if filtered_df.empty:
            customtkinter.CTkLabel(self.list_container, text=f"No sales found for {title.lower()}.").grid(row=0, column=0, pady=10)
            customtkinter.CTkButton(self.list_container, text="Back", command=self.show_graph_options).grid(row=1, column=0, pady=6)
            return

        grouped = filtered_df.groupby([filtered_df['date_parsed'].dt.strftime('%Y-%m-%d'), 'type'])['qty'].sum().unstack(fill_value=0)
        grouped = grouped.reindex(sorted(grouped.index), axis=0)

        chart_card = customtkinter.CTkFrame(self.list_container, corner_radius=12, fg_color="white")
        chart_card.grid(row=0, column=0, pady=12, padx=12, sticky="nsew")

        customtkinter.CTkLabel(chart_card, text=f"📊 {title}", font=("Arial", 14, "bold"), text_color="black").grid(row=0, column=0, pady=8)
        chart_canvas_holder = customtkinter.CTkFrame(chart_card, fg_color="transparent")
        chart_canvas_holder.grid(row=1, column=0, sticky="nsew", padx=10, pady=6)
        
        fig, ax = plt.subplots(figsize=(7, 3), dpi=100)
        grouped.plot(kind='bar', ax=ax, width=0.8, color=['#e76f51', '#f4a261'])
        ax.set_title(title)
        ax.set_ylabel("Quantity")
        ax.set_xlabel("Date")
        ax.legend(title="")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=chart_canvas_holder)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.grid(row=0, column=0, sticky="nsew")
        
        customtkinter.CTkButton(self.list_container, text="Hide Graph", command=self.clear_list_area).grid(row=2, column=0, pady=6)
        customtkinter.CTkButton(self.list_container, text="Back to Options", command=self.show_graph_options).grid(row=3, column=0, pady=6)


def create_add_stock_frame(parent_frame, app_instance):
    add_stock_frame = customtkinter.CTkFrame(parent_frame)
    add_stock_frame.grid_columnconfigure(0, weight=1)
    add_stock_frame.grid_rowconfigure(1, weight=1)

    #  New function to reset the frame to its initial state
    def reset_frame():
        # Clear the main input field and its message label
        unique_id_entry.delete(0, 'end')
        scan_msg_label.configure(text="")
        
        # Hide the details form and show the initial scan/input section
        details_section.grid_forget()
        input_section.grid(row=0, column=0, pady=20, padx=20, sticky="n")
        unique_id_entry.focus_set()


    def show_details_form(item_code, is_new_item=False, item_details=None):
        input_section.grid_forget()
        details_section.grid(row=0, column=0, pady=20, padx=20, sticky="nsew")
        
        form_widgets['name_entry'].delete(0, 'end')
        form_widgets['size_entry'].delete(0, 'end')
        form_widgets['color_entry'].delete(0, 'end')
        form_widgets['article_entry'].delete(0, 'end')
        form_widgets['quantity_entry'].delete(0, 'end')
        
        form_widgets['item_code_label'].configure(text=f"Item Code: {item_code}")
        form_widgets['item_code_label'].pack(pady=(10, 5))
        
        if is_new_item:
            customtkinter.CTkLabel(details_section, text="New Item - Details are auto-filled from ID", font=("Arial", 12)).pack(pady=(0, 5))
            
            parts = item_code.split('-')
            
            name = ""
            article = ""
            size = ""
            color = ""
            
            if len(parts) >= 4:
                name = parts[0].replace('_', ' ')
                article = parts[1].replace('_', ' ')
                size = parts[2]
                color = parts[3].replace('_', ' ')
            else:
                sg_msgbox("Error", "Invalid unique ID format. Please use BRAND-ARTICLE-SIZE-COLOR.")
                reset_frame() # Go back if format is wrong
                return

            form_widgets['name_entry'].insert(0, name)
            form_widgets['name_entry'].pack(pady=4)
            form_widgets['name_entry'].configure(state="disabled")
            
            form_widgets['size_entry'].insert(0, size)
            form_widgets['size_entry'].pack(pady=4)
            form_widgets['size_entry'].configure(state="disabled")
            
            form_widgets['color_entry'].insert(0, color)
            form_widgets['color_entry'].pack(pady=4)
            form_widgets['color_entry'].configure(state="disabled")

            form_widgets['article_entry'].insert(0, article)
            form_widgets['article_entry'].pack(pady=4)
            form_widgets['article_entry'].configure(state="disabled")

            form_widgets['quantity_entry'].configure(state="normal")
            form_widgets['quantity_entry'].pack(pady=4)
            add_stock_button.configure(command=lambda: save_new_item(item_code, name, size, color, article))
        else:
            customtkinter.CTkLabel(details_section, text="Existing Item - Details are pre-filled", font=("Arial", 12)).pack(pady=(0, 5))
            form_widgets['name_entry'].insert(0, item_details['name'])
            form_widgets['name_entry'].pack(pady=4)
            form_widgets['name_entry'].configure(state="disabled")
            
            form_widgets['size_entry'].insert(0, item_details['size'])
            form_widgets['size_entry'].pack(pady=4)
            form_widgets['size_entry'].configure(state="disabled")
            
            form_widgets['color_entry'].insert(0, item_details['color'])
            form_widgets['color_entry'].pack(pady=4)
            form_widgets['color_entry'].configure(state="disabled")

            form_widgets['article_entry'].insert(0, item_details['article'])
            form_widgets['article_entry'].pack(pady=4)
            form_widgets['article_entry'].configure(state="disabled")

            form_widgets['quantity_entry'].configure(state="normal")
            form_widgets['quantity_entry'].pack(pady=4)
            add_stock_button.configure(command=lambda: save_existing_item(item_code, item_details))

        add_stock_button.pack(pady=10)
        cancel_button.pack()
    
    def handle_ok_button(event=None):
        item_code = unique_id_entry.get().strip()
        if not item_code:
            sg_msgbox("Error", "Please enter or scan an item code.")
            return

        scan_msg_label.configure(text=f"✓ Box successfully scanned: {item_code}", text_color="green")

        df_stock = read_stock()
        if item_code in df_stock['unique_id'].values:
            item_details = df_stock[df_stock['unique_id'] == item_code].iloc[0]
            show_details_form(item_code, is_new_item=False, item_details=item_details)
        else:
            show_details_form(item_code, is_new_item=True)

    def save_new_item(item_code, name, size, color, article):
        if not form_widgets['quantity_entry'].get():
            sg_msgbox("Error", "Quantity is required!")
            return
        
        try:
            quantity = int(form_widgets['quantity_entry'].get())
            if quantity <= 0:
                sg_msgbox("Error", "Quantity must be at least 1")
                return
        except ValueError:
            sg_msgbox("Error", "Quantity must be a number")
            return

        def on_task_complete(result):
            title, msg = result
            sg_msgbox(title, msg)
            app_instance.refresh_all_data()
            reset_frame() #  Reset the frame after completion

        item_id = item_code.split('-')[0]
        details = {
            'name': name, 'size': size,
            'color': color, 'article': article,
            'quantity': quantity
        }
        run_in_thread(process_stock_action, on_task_complete, item_id, "add", item_details=details)
        
    def save_existing_item(item_code, item_details):
        if not form_widgets['quantity_entry'].get():
            sg_msgbox("Error", "Quantity is required!")
            return
        
        try:
            quantity = int(form_widgets['quantity_entry'].get())
            if quantity <= 0:
                sg_msgbox("Error", "Quantity must be at least 1")
                return
        except ValueError:
            sg_msgbox("Error", "Quantity must be a number")
            return

        def on_task_complete(result):
            title, msg = result
            sg_msgbox(title, msg)
            app_instance.refresh_all_data()
            reset_frame() #  Reset the frame after completion

        details = {
            'name': item_details['name'], 'size': item_details['size'],
            'color': item_details['color'], 'article': item_details['article'],
            'quantity': quantity
        }
        run_in_thread(process_stock_action, on_task_complete, item_details['item_id'], "add", item_details=details)

    input_section = customtkinter.CTkFrame(add_stock_frame, fg_color="transparent")
    input_section.grid(row=0, column=0, pady=20, padx=20, sticky="n")
    
    customtkinter.CTkLabel(input_section, text="➕ Add Stock Mode", font=("Arial", 20, "bold")).pack(pady=10)
    customtkinter.CTkLabel(input_section, text="Scan or type manually", font=("Arial", 14)).pack(pady=5)
    
    unique_id_entry = customtkinter.CTkEntry(input_section, placeholder_text="Scan / Enter Item Code", width=380)
    unique_id_entry.pack(pady=8, padx=20)
    
    unique_id_entry.bind('<Return>', handle_ok_button)

    scan_msg_label = customtkinter.CTkLabel(input_section, text="", font=("Arial", 12))
    scan_msg_label.pack(pady=4)

    def scan_mode():
        scan_msg_label.configure(text="Ready to scan...", text_color="blue")
        unique_id_entry.delete(0, 'end') 
        unique_id_entry.focus_set()

    def manual_mode():
        scan_msg_label.configure(text="Manual entry mode", text_color="orange")
        unique_id_entry.delete(0, 'end') 
        unique_id_entry.focus_set()
        
    button_frame = customtkinter.CTkFrame(input_section, fg_color="transparent")
    button_frame.pack(pady=8)
    
    scan_button = customtkinter.CTkButton(button_frame, text="Scan", command=scan_mode)
    scan_button.pack(side="left", padx=8)
    
    manual_button = customtkinter.CTkButton(button_frame, text="Manual", command=manual_mode)
    manual_button.pack(side="left", padx=8)
    
    ok_button = customtkinter.CTkButton(input_section, text="OK", command=handle_ok_button)
    ok_button.pack(pady=14)

    details_section = customtkinter.CTkFrame(add_stock_frame)
    
    form_widgets = {}
    form_widgets['item_code_label'] = customtkinter.CTkLabel(details_section, text="", font=("Arial", 16, "bold"))
    form_widgets['name_entry'] = customtkinter.CTkEntry(details_section, placeholder_text="Name")
    form_widgets['size_entry'] = customtkinter.CTkEntry(details_section, placeholder_text="Size")
    form_widgets['color_entry'] = customtkinter.CTkEntry(details_section, placeholder_text="Color")
    form_widgets['article_entry'] = customtkinter.CTkEntry(details_section, placeholder_text="Article")
    form_widgets['quantity_entry'] = customtkinter.CTkEntry(details_section, placeholder_text="Quantity")
    
    add_stock_button = customtkinter.CTkButton(details_section, text="✅ Add to Stock")
    
    # Changed the cancel button to call the new reset_frame function
    cancel_button = customtkinter.CTkButton(details_section, text="❌ Cancel", command=reset_frame)
    
    add_stock_button.pack(pady=10)
    cancel_button.pack()
    
    return add_stock_frame


def create_sell_item_frame(parent_frame, app_instance):
    sell_item_frame = customtkinter.CTkFrame(parent_frame)
    sell_item_frame.grid_columnconfigure(0, weight=1)
    
    # New label to display messages
    message_label = customtkinter.CTkLabel(sell_item_frame, text="", font=("Arial", 14, "bold"))
    message_label.grid(row=9, column=0, pady=10)
    
    def proceed_callback(result):
        title, msg = result
        message_label.configure(text=msg, text_color="green" if title == "Item Sold" else "red")
        app_instance.refresh_all_data()
        unique_id_entry.delete(0, 'end')
        qty_entry.delete(0, 'end')
        item_details_label.configure(text="")
        unique_id_entry.focus_set()

    def proceed(event=None):
        code = unique_id_entry.get().strip()
        code = code.replace(' ', '-')
        if not code:
            message_label.configure(text="Please enter or scan item code", text_color="red")
            return
        try:
            qty = int(qty_entry.get().strip() or "1")
            if qty <= 0:
                message_label.configure(text="Quantity must be at least 1", text_color="red")
                return
        except ValueError:
            message_label.configure(text="Quantity must be a number", text_color="red")
            return

        run_in_thread(
            process_stock_action,
            proceed_callback,
            code,
            "sell",
            item_details={'floor': 'All Floors'},
            qty=qty
        )

    def scan_mode():
        scan_msg_label.configure(text="Ready to Scan...", text_color="green")
        unique_id_entry.focus_set()
        message_label.configure(text="") # Clear message on new scan

    def manual_mode():
        scan_msg_label.configure(text="Manual entry mode", text_color="orange")
        unique_id_entry.focus_set()
        message_label.configure(text="") # Clear message on new entry

    def find_item(event=None):
        item_id = unique_id_entry.get().strip()
        item_id = item_id.replace(' ', '-')
        if not item_id:
            message_label.configure(text="Please enter or scan item code", text_color="red")
            return
        
        df = read_stock()
        if item_id in df['unique_id'].values:
            item_data = df[df['unique_id'] == item_id].iloc[0]
            item_details_label.configure(text=f"Item: {item_data['name']} | Size: {item_data['size']} | Qty in Stock: {int(item_data['quantity'])}")
            item_details_label.grid()
            qty_entry.focus_set()
            message_label.configure(text="") # Clear message on success
        else:
            message_label.configure(text=f"{item_id} not found in stock.", text_color="red")
            item_details_label.grid_remove()

    customtkinter.CTkLabel(sell_item_frame, text="➖ Sell Mode", font=("Arial", 20, "bold")).grid(row=0, column=0, pady=16)
    customtkinter.CTkLabel(sell_item_frame, text="Scan or type manually", font=("Arial", 14)).grid(row=1, column=0, pady=6)

    unique_id_entry = customtkinter.CTkEntry(sell_item_frame, placeholder_text="Scan / Enter Unique Item Code", width=380)
    unique_id_entry.grid(row=2, column=0, pady=6, padx=20)
    
    unique_id_entry.bind('<Return>', find_item) 

    qty_entry = customtkinter.CTkEntry(sell_item_frame, placeholder_text="Quantity (default 1)")
    qty_entry.grid(row=3, column=0, pady=6, padx=20)

    scan_msg_label = customtkinter.CTkLabel(sell_item_frame, text="", font=("Arial", 12))
    scan_msg_label.grid(row=4, column=0, pady=4)

    item_details_label = customtkinter.CTkLabel(sell_item_frame, text="", font=("Arial", 12, "bold"))
    item_details_label.grid(row=5, column=0, pady=6)
    item_details_label.grid_remove()

    button_frame = customtkinter.CTkFrame(sell_item_frame, fg_color="transparent")
    button_frame.grid(row=6, column=0, pady=8)

    scan_button = customtkinter.CTkButton(button_frame, text="Scan", command=lambda: scan_mode())
    scan_button.pack(side="left", padx=8)

    manual_button = customtkinter.CTkButton(button_frame, text="Manual", command=lambda: manual_mode())
    manual_button.pack(side="left", padx=8)
    
    find_button = customtkinter.CTkButton(sell_item_frame, text="🔍 Find Item", command=find_item)
    find_button.grid(row=7, column=0, pady=8)

    customtkinter.CTkButton(sell_item_frame, text="✅ Sell Item", command=proceed).grid(row=8, column=0, pady=12)

    return sell_item_frame


def create_return_item_frame(parent_frame, app_instance):
    return_item_frame = customtkinter.CTkFrame(parent_frame)
    return_item_frame.grid_columnconfigure(0, weight=1)

    def return_now_callback(result):
        if isinstance(result, tuple) and len(result) == 3:
            title, stock_before, stock_after = result
            msg = f"Item successfully returned!\n\nStock before return: {stock_before}\nStock after return: {stock_after}"
            sg_msgbox(title, msg)
        else:
            title, msg = result
            sg_msgbox(title, msg)

        app_instance.refresh_all_data()

        unique_id_entry.delete(0, 'end')
        qty_entry.delete(0, 'end')
        item_details_label.grid_remove()
        scan_msg_label.configure(text="")
        unique_id_entry.focus_set()
    
    def return_now(event=None):
        item_id = unique_id_entry.get().strip()
        if not item_id:
            sg_msgbox("Error", "Please enter or scan an item code.")
            return
        
        df_stock = read_stock()
        if item_id not in df_stock['unique_id'].values:
            sg_msgbox("Error", f"{item_id} not found in stock.")
            return
        item_details = df_stock[df_stock['unique_id'] == item_id].iloc[0].to_dict()

        try:
            qty = int(qty_entry.get().strip() or "1")
            if qty <= 0:
                sg_msgbox("Error", "Quantity must be at least 1.")
                return
        except ValueError:
            sg_msgbox("Error", "Quantity must be a number.")
            return

        run_in_thread(
            process_stock_action,
            return_now_callback,
            item_id,
            "return",
            item_details=item_details,
            qty=qty
        )

    def scan_mode():
        scan_msg_label.configure(text="Ready to Scan...", text_color="green")
        unique_id_entry.focus_set()

    def manual_mode():
        scan_msg_label.configure(text="Manual entry mode", text_color="orange")
        unique_id_entry.focus_set()
        
    def find_item(event=None):
        item_id = unique_id_entry.get().strip()
        if not item_id:
            sg_msgbox("Error", "Please enter or scan item code")
            return
        
        df = read_stock()
        if item_id in df['unique_id'].values:
            item_data = df[df['unique_id'] == item_id].iloc[0]
            item_details_label.configure(text=f"Item: {item_data['name']} | Size: {item_data['size']} | Qty in Stock: {int(item_data['quantity'])}")
            item_details_label.grid()
            qty_entry.focus_set()
        else:
            sg_msgbox("Error", f"{item_id} not found in stock.")
            item_details_label.grid_remove()

    customtkinter.CTkLabel(return_item_frame, text="↩ Return / Restock Item", font=("Arial", 16, "bold")).grid(row=0, column=0, pady=(12, 6))
    customtkinter.CTkLabel(return_item_frame, text="Scan or type manually", font=("Arial", 14)).grid(row=1, column=0, pady=6)

    unique_id_entry = customtkinter.CTkEntry(return_item_frame, placeholder_text="Enter Item Unique Code", width=380)
    unique_id_entry.grid(row=2, column=0, pady=(6, 10))
    
    unique_id_entry.bind('<Return>', find_item)

    qty_entry = customtkinter.CTkEntry(return_item_frame, placeholder_text="Enter Quantity to Return")
    qty_entry.grid(row=3, column=0, pady=(6, 12))

    scan_msg_label = customtkinter.CTkLabel(return_item_frame, text="", font=("Arial", 12))
    scan_msg_label.grid(row=4, column=0, pady=4)
    
    item_details_label = customtkinter.CTkLabel(return_item_frame, text="", font=("Arial", 12, "bold"))
    item_details_label.grid(row=5, column=0, pady=6)
    item_details_label.grid_remove()

    button_frame = customtkinter.CTkFrame(return_item_frame, fg_color="transparent")
    button_frame.grid(row=6, column=0, pady=8)

    scan_button = customtkinter.CTkButton(button_frame, text="Scan", command=lambda: scan_mode())
    scan_button.pack(side="left", padx=8)

    manual_button = customtkinter.CTkButton(button_frame, text="Manual", command=lambda: manual_mode())
    manual_button.pack(side="left", padx=8)
    
    find_button = customtkinter.CTkButton(return_item_frame, text="🔍 Find Item", command=find_item)
    find_button.grid(row=7, column=0, pady=8)

    customtkinter.CTkButton(return_item_frame, text="↩ Return Item", command=return_now).grid(row=8, column=0, pady=12)

    return return_item_frame


def create_sales_history_frame(parent_frame, app_instance):
    sales_history_frame = customtkinter.CTkFrame(parent_frame)
    sales_history_frame.grid_columnconfigure(0, weight=1)
    sales_history_frame.grid_rowconfigure(1, weight=1)

    main_frame = customtkinter.CTkFrame(sales_history_frame)
    main_frame.grid(row=0, column=0, pady=10, padx=20, sticky="nsew")
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_rowconfigure(2, weight=1)

    def show_sales_report(date_str=None, floor_filter=None):
        for widget in result_frame.winfo_children():
            widget.destroy()

        df = read_sales()
        if df.empty:
            customtkinter.CTkLabel(result_frame, text="No sales recorded yet.").grid(row=0, column=0, pady=20)
            return

        filtered_df = df
        if date_str and date_str != "all":
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
                filtered_df = filtered_df[filtered_df['date'] == date_str]
            except ValueError:
                sg_msgbox("Invalid Date", "Please enter the date in YYYY-MM-DD format.")
                return

        if floor_filter and floor_filter != "All Floors":
            filtered_df = filtered_df[filtered_df['floor'] == floor_filter]

        if filtered_df.empty:
            customtkinter.CTkLabel(result_frame, text="No sales found for the selected criteria.").grid(row=0, column=0, pady=20)
        else:
            headers = ["Date", "Time", "Name", "Size", "Qty", "Type", "Floor"]
            for i, h in enumerate(headers):
                customtkinter.CTkLabel(result_frame, text=h, font=("Arial", 12, "bold")).grid(row=0, column=i, padx=5, pady=5)

            for i, (_, row) in enumerate(filtered_df.iterrows()):
                data = [row['date'], row['time'], row['name'], row['size'], int(row['qty']), row['type'], row['floor']]
                for j, val in enumerate(data):
                    customtkinter.CTkLabel(result_frame, text=str(val)).grid(row=i+1, column=j, padx=5, pady=3, sticky="w")
    
    customtkinter.CTkLabel(main_frame, text="Sales Report", font=("Arial", 16, "bold")).grid(row=0, column=0, pady=(10, 6))
    
    filter_frame = customtkinter.CTkFrame(main_frame)
    filter_frame.grid(row=1, column=0, pady=6, padx=10, sticky="ew")
    filter_frame.grid_columnconfigure(1, weight=1)
    filter_frame.grid_columnconfigure(3, weight=1)
    
    customtkinter.CTkLabel(filter_frame, text="Enter Date (YYYY-MM-DD):", font=("Arial", 12)).grid(row=0, column=0, padx=(10, 6), sticky="w")
    date_entry = customtkinter.CTkEntry(filter_frame, placeholder_text="e.g., 2025-09-16")
    date_entry.grid(row=0, column=1, padx=6, sticky="ew")
    
    customtkinter.CTkLabel(filter_frame, text="Filter by Floor:", font=("Arial", 12)).grid(row=0, column=2, padx=(20, 6), sticky="w")
    floor_options = ["All Floors", "Floor 1", "Floor 2"]
    floor_var = customtkinter.StringVar(value=floor_options[0])
    floor_menu = customtkinter.CTkOptionMenu(filter_frame, values=floor_options, variable=floor_var)
    floor_menu.grid(row=0, column=3, padx=6, sticky="ew")
    
    button_frame = customtkinter.CTkFrame(filter_frame, fg_color="transparent")
    button_frame.grid(row=1, column=0, columnspan=4, pady=(10, 0))
    
    show_button = customtkinter.CTkButton(button_frame, text="Show Report", command=lambda: show_sales_report(date_entry.get().strip(), floor_var.get()))
    show_button.pack(side="left", padx=5)

    show_all_button = customtkinter.CTkButton(button_frame, text="Show All Sales", command=lambda: show_sales_report(date_str="all", floor_filter=floor_var.get()))
    show_all_button.pack(side="left", padx=5)

    result_frame = customtkinter.CTkScrollableFrame(main_frame)
    result_frame.grid(row=2, column=0, pady=10, padx=10, sticky="nsew")
    result_frame.grid_columnconfigure(0, weight=1)
    
    show_sales_report(datetime.now().strftime("%Y-%m-%d"))

    return sales_history_frame


def create_search_inventory_frame(parent_frame, app_instance):
    search_inventory_frame = customtkinter.CTkFrame(parent_frame)
    search_inventory_frame.grid_columnconfigure(0, weight=1)
    search_inventory_frame.grid_rowconfigure(2, weight=1)

    def open_edit_form(item_details):
        edit_frame = customtkinter.CTkFrame(search_inventory_frame)
        edit_frame.grid(row=0, column=0, rowspan=3, pady=12, padx=20, sticky="nsew")

        for widget in search_inventory_frame.winfo_children():
            if widget != edit_frame:
                widget.grid_forget()

        customtkinter.CTkLabel(edit_frame, text=f"Edit Item: {item_details['unique_id']}", font=("Arial", 16, "bold")).pack(pady=8)

        name_entry = customtkinter.CTkEntry(edit_frame, placeholder_text="Name")
        name_entry.insert(0, item_details['name'])
        name_entry.pack(pady=4)

        size_entry = customtkinter.CTkEntry(edit_frame, placeholder_text="Size")
        size_entry.insert(0, item_details['size'])
        size_entry.pack(pady=4)

        color_entry = customtkinter.CTkEntry(edit_frame, placeholder_text="Color")
        color_entry.insert(0, item_details['color'])
        color_entry.pack(pady=4)

        article_entry = customtkinter.CTkEntry(edit_frame, placeholder_text="Article")
        article_entry.insert(0, item_details['article'])
        article_entry.pack(pady=4)
        
        quantity_entry = customtkinter.CTkEntry(edit_frame, placeholder_text="Quantity")
        quantity_entry.insert(0, str(int(item_details['quantity'])))
        quantity_entry.pack(pady=6)

        def save_changes():
            try:
                df = read_stock()
                idx = df[df['unique_id'] == item_details['unique_id']].index[0]
            except Exception:
                sg_msgbox("Error", "Couldn't find the item in stock (it may have been deleted).")
                app_instance.show_frame("SearchInventory")
                return

            new_unique_id = f"{item_details['item_id']}-{size_entry.get()}-{color_entry.get()}-{article_entry.get()}"

            if new_unique_id != item_details['unique_id'] and new_unique_id in df['unique_id'].values:
                sg_msgbox("Error", "New details create an ID that already exists. Please choose different details.")
                return

            try:
                df.at[idx, 'name'] = name_entry.get()
                df.at[idx, 'size'] = size_entry.get()
                df.at[idx, 'color'] = color_entry.get()
                df.at[idx, 'article'] = article_entry.get()
                df.at[idx, 'quantity'] = int(quantity_entry.get())
                df.at[idx, 'unique_id'] = new_unique_id
            except ValueError:
                sg_msgbox("Error", "Quantity must be a number")
                return

            write_stock(df)
            sg_msgbox("Success", "Item details have been updated!")
            app_instance.update_data()
            app_instance.show_frame("SearchInventory")

        customtkinter.CTkButton(edit_frame, text="Save Changes", command=save_changes).pack(pady=8)
        customtkinter.CTkButton(edit_frame, text="Cancel", command=lambda: app_instance.show_frame("SearchInventory")).pack(pady=4)

    def delete_item(unique_id):
        confirmed = sg_confirm("Confirm Deletion", f"Are you sure you want to delete {unique_id}?\nThis action cannot be undone.")
        if not confirmed:
            return
        df = read_stock()
        df = df[df['unique_id'] != unique_id]
        write_stock(df)
        sg_msgbox("Success", f"{unique_id} has been deleted.")
        app_instance.update_data()
        app_instance.show_frame("SearchInventory")

    def do_search(event=None):
        df = read_stock()
        q = search_entry.get().lower().strip()
        for w in result_frame.winfo_children():
            w.destroy()

        if q == "":
            customtkinter.CTkLabel(result_frame, text="Please enter a search keyword").pack(pady=20)
            return

        result = df[df.apply(lambda r: q in str(list(r.values)).lower(), axis=1)]

        if result.empty:
            customtkinter.CTkLabel(result_frame, text="No results found").pack(pady=20)
        else:
            for _, row in result.iterrows():
                result_item_frame = customtkinter.CTkFrame(result_frame, fg_color="transparent")
                result_item_frame.pack(fill="x", pady=6)

                name_label = customtkinter.CTkLabel(result_item_frame, text=f"Name: {row['name']}", font=("Arial", 12, "bold"))
                name_label.pack(anchor="w")

                details_label = customtkinter.CTkLabel(result_item_frame, text=f"Size: {row['size']}  |  Color: {row['color']}  |  Article: {row['article']}  |  Qty: {int(row['quantity'])}")
                details_label.pack(anchor="w")

                button_frame = customtkinter.CTkFrame(result_item_frame, fg_color="transparent")
                button_frame.pack(pady=(6, 0), anchor="e")

                customtkinter.CTkButton(button_frame, text="Edit", command=lambda r=row: open_edit_form(r)).pack(side="left", padx=6)
                customtkinter.CTkButton(button_frame, text="Delete", fg_color="red", hover_color="#c70000", command=lambda unique_id=row['unique_id']: delete_item(unique_id)).pack(side="left", padx=6)
    
    customtkinter.CTkLabel(search_inventory_frame, text="🔍 Search Inventory", font=("Arial", 20, "bold")).grid(row=0, column=0, pady=14)

    search_bar_frame = customtkinter.CTkFrame(search_inventory_frame, fg_color="transparent")
    search_bar_frame.grid(row=1, column=0, padx=10, pady=8, sticky="ew")
    search_bar_frame.grid_columnconfigure(0, weight=1)

    search_entry = customtkinter.CTkEntry(search_bar_frame, placeholder_text="Enter keyword...", width=520)
    search_entry.grid(row=0, column=0, padx=(0, 5), sticky="ew")
    search_entry.bind('<Return>', do_search)

    search_button = customtkinter.CTkButton(search_bar_frame, text="Search", command=do_search)
    search_button.grid(row=0, column=1, padx=(5, 0))

    result_frame = customtkinter.CTkScrollableFrame(search_inventory_frame, width=700, height=380)
    result_frame.grid(row=2, column=0, pady=8, padx=10, sticky="nsew")

    return search_inventory_frame


def create_export_reports_frame(parent_frame):
    export_reports_frame = customtkinter.CTkFrame(parent_frame)
    export_reports_frame.grid_columnconfigure(0, weight=1)
    export_reports_frame.grid_rowconfigure((0,1,2), weight=1)

    def export_stock():
        df = read_stock()
        if df.empty:
            sg_msgbox("Export", "No stock data available to export.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path:
            df.to_excel(path, index=False)
            sg_msgbox("Export", f"Stock exported to {path}")

    def export_sales(date_str=None, floor_filter=None):
        df = read_sales()
        if df.empty:
            sg_msgbox("Export", "No sales data available to export.")
            return

        filtered_df = df
        if date_str and date_str != "all":
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                sg_msgbox("Invalid Date", "Please enter the date in YYYY-MM-DD format.")
                return
            filtered_df = df[df['date'] == date_str]

        if floor_filter and floor_filter != "All Floors":
            filtered_df = filtered_df[filtered_df['floor'] == floor_filter]

        if filtered_df.empty:
            sg_msgbox("Export", f"No sales found for the date {date_str}. Nothing to export.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path:
            filtered_df.to_excel(path, index=False)
            sg_msgbox("Export", f"Sales exported to {path}")

    customtkinter.CTkLabel(export_reports_frame, text="📥 Export Reports", font=("Arial", 20, "bold")).grid(row=0, column=0, pady=18, sticky="n")

    stock_frame = customtkinter.CTkFrame(export_reports_frame, corner_radius=12)
    stock_frame.grid(row=1, column=0, pady=8, padx=20, sticky="nsew")
    stock_frame.grid_columnconfigure(0, weight=1)
    customtkinter.CTkLabel(stock_frame, text="Export Stock Report", font=("Arial", 14, "bold")).grid(row=0, column=0, pady=6)
    customtkinter.CTkButton(stock_frame, text="Export Stock", command=export_stock).grid(row=1, column=0, pady=8)

    sales_frame = customtkinter.CTkFrame(export_reports_frame, corner_radius=12)
    sales_frame.grid(row=2, column=0, pady=8, padx=20, sticky="nsew")
    sales_frame.grid_columnconfigure(0, weight=1)
    customtkinter.CTkLabel(sales_frame, text="Export Sales Report by Date", font=("Arial", 14, "bold")).grid(row=0, column=0, pady=6)

    date_frame = customtkinter.CTkFrame(sales_frame, fg_color="transparent")
    date_frame.grid(row=1, column=0, pady=6, padx=10)
    date_frame.grid_columnconfigure(1, weight=1)
    date_frame.grid_columnconfigure(3, weight=1)
    
    customtkinter.CTkLabel(date_frame, text="Date (YYYY-MM-DD):").grid(row=0, column=0, padx=6)
    date_entry = customtkinter.CTkEntry(date_frame, placeholder_text="e.g., 2025-09-16")
    date_entry.grid(row=0, column=1, padx=6)
    
    customtkinter.CTkLabel(date_frame, text="Filter by Floor:").grid(row=0, column=2, padx=(20, 6))
    floor_options = ["All Floors", "Floor 1", "Floor 2"]
    floor_var = customtkinter.StringVar(value=floor_options[0])
    floor_menu = customtkinter.CTkOptionMenu(date_frame, values=floor_options, variable=floor_var)
    floor_menu.grid(row=0, column=3, padx=6)


    customtkinter.CTkButton(sales_frame, text="Export Sales for this Date", command=lambda: export_sales(date_entry.get().strip(), floor_var.get())).grid(row=2, column=0, pady=8)
    customtkinter.CTkButton(sales_frame, text="Export All Sales", command=lambda: export_sales(date_str="all", floor_filter=floor_var.get())).grid(row=3, column=0, pady=4)

    return export_reports_frame


# --- QR Code Generator Frame ---
def create_qr_generator_frame(parent_frame):
    qr_frame = customtkinter.CTkFrame(parent_frame)
    qr_frame.grid_columnconfigure(0, weight=1)
    
    def generate_qr_code_from_details():
        brand = brand_entry.get().strip().replace(' ', '-')
        size = size_entry.get().strip()
        color = color_entry.get().strip().replace(' ', '-')
        article = article_entry.get().strip().replace(' ', '-')

        if not all([brand, size, color, article]):
            sg_msgbox("Error", "Please fill in all the details.")
            return

        unique_id = f"{brand.upper()}-{article.upper()}-{size}-{color.upper()}"
        
        qr = qrcode.QRCode(version=1, box_size=10, border=4, error_correction=qrcode.constants.ERROR_CORRECT_L)
        qr.add_data(unique_id)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        
        img_tk = ImageTk.PhotoImage(img_qr)
        qr_label.configure(image=img_tk, text="")
        qr_label.image = img_tk
        
        sg_msgbox("QR Code Generated", f"Unique ID: {unique_id}")

    def save_qr_code():
        brand = brand_entry.get().strip()
        size = size_entry.get().strip()
        color = color_entry.get().strip()
        article = article_entry.get().strip()
        if not all([brand, size, color, article]):
            sg_msgbox("Error", "Please generate a QR code first.")
            return

        unique_id = f"{brand.upper().replace(' ', '-')}-{article.upper().replace(' ', '-')}-{size}-{color.upper().replace(' ', '-')}"
        qr = qrcode.QRCode(version=1, box_size=10, border=4, error_correction=qrcode.constants.ERROR_CORRECT_L)
        qr.add_data(unique_id)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        
        path = filedialog.asksaveasfilename(defaultextension=".png", initialfile=f"QR_{unique_id}.png", filetypes=[("PNG files", "*.png")])
        if path:
            img_qr.save(path)
            sg_msgbox("Success", f"QR code saved to {path}")

    customtkinter.CTkLabel(qr_frame, text="Generate QR Code", font=("Arial", 20, "bold")).pack(pady=10)
    customtkinter.CTkLabel(qr_frame, text="Enter Item Details to Create QR Code", font=("Arial", 14)).pack(pady=5)
    
    brand_entry = customtkinter.CTkEntry(qr_frame, placeholder_text="Brand Name")
    brand_entry.pack(pady=5)
    size_entry = customtkinter.CTkEntry(qr_frame, placeholder_text="Size")
    size_entry.pack(pady=5)
    color_entry = customtkinter.CTkEntry(qr_frame, placeholder_text="Color")
    color_entry.pack(pady=5)
    article_entry = customtkinter.CTkEntry(qr_frame, placeholder_text="Article")
    article_entry.pack(pady=5)

    customtkinter.CTkButton(qr_frame, text="Generate QR Code", command=generate_qr_code_from_details).pack(pady=10)

    qr_display_frame = customtkinter.CTkFrame(qr_frame, width=200, height=200, fg_color="white", corner_radius=12)
    qr_display_frame.pack(pady=10)
    qr_label = customtkinter.CTkLabel(qr_display_frame, text="QR Code will appear here", font=("Arial", 12), width=200, height=200)
    qr_label.pack()

    customtkinter.CTkButton(qr_frame, text="Save QR Code as PNG", command=save_qr_code).pack(pady=5)

    return qr_frame


# ---- App class ----
class App(customtkinter.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Shoe Store Manager")
        try:
            self.state("zoomed")
        except Exception:
            self.geometry("1200x700")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.frames = {}
        self.sidebar_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=10, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1)
        
        # Define start_row before the try block
        start_row = 1 # Set a default value

        # --- CODE FOR THE LOGO ---
        try:
            from PIL import Image
            # Use the absolute path to the logo file.
            logo_path = r"C:\Users\sawan\OneDrive\Desktop\SHOEapp\logo.jpg"
            self.logo_image = customtkinter.CTkImage(
                light_image=Image.open(logo_path),
                dark_image=Image.open(logo_path),
                size=(120, 120)
            )
            self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="", image=self.logo_image)
            self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
            
            # The row numbers for all buttons have been shifted down by 1.
            start_row = 1
        except FileNotFoundError:
            print("Warning: logo.jpg not found. Skipping logo display. Ensure the file is at the correct path.")
            self.sidebar_title = customtkinter.CTkLabel(self.sidebar_frame, text="Navigation", font=("Arial", 16, "bold"))
            self.sidebar_title.grid(row=0, column=0, padx=14, pady=18)
            # start_row remains 1 from the default value
            start_row = 1
            
        self.dashboard_button = customtkinter.CTkButton(self.sidebar_frame, text="🏠 Dashboard", command=lambda: self.show_frame("Dashboard"))
        self.dashboard_button.grid(row=start_row, column=0, padx=12, pady=6)

        self.add_stock_button = customtkinter.CTkButton(self.sidebar_frame, text="➕ Add Stock", command=lambda: self.show_frame("AddStock"))
        self.add_stock_button.grid(row=start_row + 1, column=0, padx=12, pady=6)

        self.sell_item_button = customtkinter.CTkButton(self.sidebar_frame, text="➖ Sell Item", command=lambda: self.show_frame("SellItem"))
        self.sell_item_button.grid(row=start_row + 2, column=0, padx=12, pady=6)

        self.return_item_button = customtkinter.CTkButton(self.sidebar_frame, text="↩ Return Item", command=lambda: self.show_frame("ReturnItem"))
        self.return_item_button.grid(row=start_row + 3, column=0, padx=12, pady=6)

        self.sales_history_button = customtkinter.CTkButton(self.sidebar_frame, text="📊 Sales History", command=lambda: self.show_frame("SalesHistory"))
        self.sales_history_button.grid(row=start_row + 4, column=0, padx=12, pady=6)

        self.search_inventory_button = customtkinter.CTkButton(self.sidebar_frame, text="🔍 Search Inventory", command=lambda: self.show_frame("SearchInventory"))
        self.search_inventory_button.grid(row=start_row + 5, column=0, padx=12, pady=6)

        self.export_reports_button = customtkinter.CTkButton(self.sidebar_frame, text="📥 Export Reports", command=lambda: self.show_frame("ExportReports"))
        self.export_reports_button.grid(row=start_row + 6, column=0, padx=12, pady=6)

        self.qr_generator_button = customtkinter.CTkButton(self.sidebar_frame, text="Create QR Code", command=lambda: self.show_frame("QRCodeGenerator"))
        self.qr_generator_button.grid(row=start_row + 7, column=0, padx=12, pady=6)

        self.refresh_button = customtkinter.CTkButton(self.sidebar_frame, text="🔄 Refresh", command=lambda: self.refresh_all_data(show_msg=True))
        self.refresh_button.grid(row=start_row + 8, column=0, padx=12, pady=6, sticky="s")
        self.sidebar_frame.grid_rowconfigure(start_row + 8, weight=1)
        
        self.frames["Dashboard"] = DashboardFrame(self, self)
        self.frames["AddStock"] = create_add_stock_frame(self, self)
        self.frames["SellItem"] = create_sell_item_frame(self, self)
        self.frames["ReturnItem"] = create_return_item_frame(self, self)
        self.frames["SalesHistory"] = create_sales_history_frame(self, self)
        self.frames["SearchInventory"] = create_search_inventory_frame(self, self)
        self.frames["ExportReports"] = create_export_reports_frame(self)
        self.frames["QRCodeGenerator"] = create_qr_generator_frame(self)
        
        self.current_frame = None
        self.show_frame("Dashboard")

    def show_frame(self, frame_name):
        if self.current_frame:
            self.current_frame.grid_forget()

        frame = self.frames[frame_name]
        frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.current_frame = frame

        if frame_name == "Dashboard":
            self.frames["Dashboard"].update_data()
        elif hasattr(frame, 'show_sales_report'):
            frame.show_sales_report()
            
    def refresh_all_data(self, show_msg=False):
        current_frame_name = None
        for name, frame in self.frames.items():
            if frame == self.current_frame:
                current_frame_name = name
                break
            
        if "Dashboard" in self.frames:
            self.frames["Dashboard"].update_data()

        if current_frame_name == "SalesHistory":
            self.frames["SalesHistory"].show_sales_report(datetime.now().strftime("%Y-%m-%d"))
        
        if show_msg:
            sg_msgbox("Refresh Successful", "Data has been refreshed!")

if __name__ == "__main__":
    app = App()
    app.mainloop()