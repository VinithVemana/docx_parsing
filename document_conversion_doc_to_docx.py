import random
import os
import shutil
from subprocess import CalledProcessError, check_output, Popen, PIPE
import multiprocessing as mp
import glob
import csv
import json

def check_kill_process(pattern: str):
    try:
        cmd = f"pgrep -f '{pattern}'"
        proc = Popen(cmd, stdout=PIPE, shell=True)
        out, _ = proc.communicate()
    except Exception:
        return "Failed"
    
    out = out.decode("utf-8").strip()
    if out:
        print(f"killing process {pattern}", end='')
        cmd = f"pkill -f '{pattern}'"
        os.system(cmd)
        print(" - Finished")
        return "Killed"
    return "Failed"

def _build_cmd(cmd: str, **args: dict):
    for search, replace in args.items():
        cmd = cmd.replace('{' + str(search) + '}', str(replace))
    cmd = cmd.replace("'", '"')
    return cmd

def _run_cmd(cmd: str, timeout: int):
    print('command for convert:', cmd)
    output = None
    try:
        output = check_output(cmd, timeout=timeout, shell=True)
    except CalledProcessError as e:
        print('Error Command:', e)
        output = e.output
        return False, output
    return True, output

def clean_temp_profiles():
    temp_profile_pattern = os.path.expanduser("~/libreOffice1/conversion_*_profile")
    temp_profiles = glob.glob(temp_profile_pattern)
    for profile in temp_profiles:
        try:
            shutil.rmtree(profile)
            print(f"Removed temporary profile directory: {profile}")
        except Exception as cleanup_error:
            print(f"Failed to remove temporary profile directory: {cleanup_error}")

def doc_to_docx(source: str, target: str, app: str, export: bool = False, doc_type: str = 'normal'):
     #app="/usr/bin/soffice"
    cmd_template = "{app} --headless --norestore --invisible --nocrashreport --nodefault --nologo --nofirststartwizard --norestore {user_profile} --convert-to docx --outdir {target_dir} {source}"
    
    temp_user_profile = None
    for _ in range(10):
        user_profile_dir = os.path.expanduser(f'~/libreOffice1/conversion_{random.randint(0, 10000000)}_profile')
        user_profile = f"-env:UserInstallation=file://{user_profile_dir}"
        if not os.path.exists(user_profile_dir):
            os.makedirs(user_profile_dir)
            temp_user_profile = user_profile_dir
            break
    
    target_dir = os.path.dirname(target)
    cmd = _build_cmd(cmd_template, app=app, user_profile=user_profile, target_dir=target_dir, source=source)
    
    file_size = os.path.getsize(source)
    post_process_secs = 30
    # extra_time = {'normal': 180, 'medium': 720, 'large': 2160}.get(doc_type, 180) - post_process_secs
    extra_time = {'normal': 480, 'medium': 720, 'large': 2160}.get(doc_type, 480) - post_process_secs

    print(f"Got file size {file_size / 1024}kb, waiting time extra {extra_time / 60} min for doc type {doc_type}")
    
    try:
        success, cmd_output = _run_cmd(cmd, timeout=extra_time)
        cmd_output = cmd_output.decode("utf-8").strip() if cmd_output else ""
    except Exception as e:
        if temp_user_profile and os.path.exists(temp_user_profile):
            shutil.rmtree(temp_user_profile)
        check_kill_process(pattern=temp_user_profile)
        return False, str(e)
    
    if temp_user_profile and os.path.exists(temp_user_profile):
        try:
            shutil.rmtree(temp_user_profile)
            print(f"Removed temporary profile directory: {temp_user_profile}")
        except Exception as cleanup_error:
            print(f"Failed to remove temporary profile directory: {cleanup_error}")
    
    if success:
        target_path, _ = os.path.splitext(target)
        source_docx = target_path + ".docx"
        target_readable_docx = target_path + "_readable.docx"
        cmd = f'mv {source_docx} {target_readable_docx}'
        if export:
            return True, source_docx
        success, _ = _run_cmd(cmd, timeout=120)
        return success, cmd_output
    
    print("Doc2Docx cmd Status: Failed")
    return False, cmd_output

def get_already_converted_files(target_dir):
    """Get a set of base filenames (without extension) that have already been converted."""
    already_converted = set()
    if os.path.exists(target_dir):
        for f in os.listdir(target_dir):
            if f.endswith('.docx'):
                # Strip the .docx extension to get the base name
                base_name = os.path.splitext(f)[0]
                already_converted.add(base_name)
    return already_converted

def convert_file(file_path, target_dir, failed_conversions):
    target_file = os.path.join(target_dir, os.path.basename(file_path).replace('.doc', '.docx'))
    status, output = doc_to_docx(source=file_path, target=target_file, app=args["libreoffice_path"], export=True, doc_type='normal')
    print(f"Conversion Status for {file_path}: {status}")
    print(f"Output for {file_path}: {output}")
    
    if not status:
        # Log failed conversion in the shared list
        failed_conversions.append(file_path)

def save_failed_conversions(failed_conversions, csv_file):
    # Write the failed conversions to a CSV file
    with open(csv_file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["File Path"])  # Header
        for file_path in failed_conversions:
            writer.writerow([file_path])

def main(args):
    source_dir = args['input_directory_name']
    target_dir = args['output_directory_name']
    csv_file = args['failed_log']
    
    os.makedirs(target_dir, exist_ok=True)
    
    # Get the set of already converted base filenames
    already_converted = get_already_converted_files(target_dir)
    
    # Filter out files that have already been converted
    doc_files = []
    skipped_count = 0
    for f in os.listdir(source_dir):
        if f.endswith('.doc'):
            base_name = os.path.splitext(f)[0]
            if base_name in already_converted:
                skipped_count += 1
                continue
            doc_files.append(os.path.join(source_dir, f))
    
    print(f"Total .doc files found: {len(doc_files) + skipped_count}")
    print(f"Already converted (skipping): {skipped_count}")
    print(f"Files to convert: {len(doc_files)}")
    
    if not doc_files:
        print("No new files to convert. Exiting.")
        return
    
    # Use a manager list to track failed conversions
    with mp.Manager() as manager:
        failed_conversions = manager.list()  # Shared list for failed conversions
        
        with mp.Pool(args['num_processes']) as pool:
            pool.starmap(convert_file, [(file, target_dir, failed_conversions) for file in doc_files])

        # Save the failed conversions to a CSV file
        save_failed_conversions(failed_conversions, csv_file)

if __name__ == "__main__":
    # Load configuration
    with open('config.json', 'r') as f:
        config = json.load(f)
    args = config['doc_to_docx_params']
    main(args)