try:
    import face_recognition
    print("Face Recognition OK")
except Exception as e:
    import traceback
    traceback.print_exc()

input("Presione Enter para salir...")