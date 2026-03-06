@app.route('/batch-upload', methods=['POST'])
def batch_upload():
    try:
        files = request.files.getlist('files')
        upload_number = request.form.get('upload_number', '').strip()
        
        if not files:
            return jsonify({'success': False, 'error': 'No files provided'})
        
        if not upload_number:
            return jsonify({'success': False, 'error': 'No upload number provided'})
        
        results = []
        
        for file in files:
            if file and file.filename.endswith('.pdf'):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                try:
                    doc_data = process_pdf(filepath, upload_number)
                    documents.append(doc_data)
                    results.append({
                        'filename': filename,
                        'success': True,
                        'data': doc_data
                    })
                except Exception as e:
                    results.append({
                        'filename': filename,
                        'success': False,
                        'error': str(e)
                    })
                
                # Clean up
                if os.path.exists(filepath):
                    os.remove(filepath)
        
        return jsonify({
            'success': True,
            'results': results
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
