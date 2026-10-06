"""
streamlit_app.py
Điểm khởi động cho Streamlit Cloud (trỏ trực tiếp vào app.py)
"""
import runpy

if __name__ == "__main__":
    runpy.run_path("app.py", run_name="__main__")
else:
    # Khi Streamlit load module
    runpy.run_path("app.py")
