import os

def delete_mp4_files(folder_path):
    """Deletes all .mp4 files in the specified folder."""
    deleted_count = 0
    for filename in os.listdir(folder_path):
        if filename.endswith(".mp4"):
            file_path = os.path.join(folder_path, filename)
            os.remove(file_path)
            deleted_count += 1
            print(f"Deleted: {filename}")
    print(f"\nFinished. {deleted_count} .mp4 files were deleted.")

if __name__ == "__main__":
    folder = input("Enter the folder path to delete .mp4 files: ").strip()
    delete_mp4_files(folder)
