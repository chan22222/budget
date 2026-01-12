#!/usr/bin/env python3
import sys
import msoffcrypto
import io

def decrypt_file(input_path, output_path, password):
    try:
        with open(input_path, 'rb') as f:
            file = msoffcrypto.OfficeFile(f)
            if file.is_encrypted():
                file.load_key(password=password)
                with open(output_path, 'wb') as out:
                    file.decrypt(out)
                print(f"SUCCESS:{output_path}")
            else:
                # 암호화 안 됨 - 그냥 복사
                with open(input_path, 'rb') as src:
                    with open(output_path, 'wb') as dst:
                        dst.write(src.read())
                print(f"SUCCESS:{output_path}")
    except Exception as e:
        print(f"ERROR:{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: decrypt.py <input> <output> <password>")
        sys.exit(1)

    decrypt_file(sys.argv[1], sys.argv[2], sys.argv[3])
