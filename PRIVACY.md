# Privacy notice

Last updated: 4 September 2026

## What the application processes

MCQ Exam Builder processes files that a user deliberately uploads, including documents, spreadsheets, text-based PDFs, text/markup files, structured question banks, and question images. These files may contain questions, answer choices, correct answers, source filenames, and embedded images.

## How processing works

The application does not use generative-AI services, external APIs, databases, analytics, or cloud-storage integrations. Uploaded content and generated files are held within the active Streamlit application session. Legacy Word, RTF, and OpenDocument text files are converted on the same server using LibreOffice Writer with macros and active content disabled; automatic link updating is disabled.

When the application is used on a hosted website, uploaded files are processed in the hosting provider's computing environment. The repository owner and hosting provider may also retain ordinary operational logs. This project does not intentionally add uploaded exam content to those logs.

## Storage and retention

The application does not intentionally create persistent storage for uploaded question banks or generated examinations. Session data may remain in server memory until the session ends, the application restarts, or the hosting platform removes it.

Document conversion uses a temporary directory containing the source, converted DOCX, and a disposable LibreOffice profile. The application removes this directory when conversion finishes or fails. Word export can also use temporary image files, removed after processing. Abrupt server termination can interrupt cleanup; these temporary files remain subject to the hosting environment's storage lifecycle. The application does not create a shared cache of exam content.

Users should download required outputs before closing the session. They should not upload personal data, confidential examinations, or restricted content unless their institution permits processing on the selected hosting platform.

## Sharing

The application does not intentionally sell or share uploaded content. Infrastructure providers process data only as required to host and operate the service, subject to their own terms and privacy notices.

## Deployment transparency

The public application is hosted on Streamlit Community Cloud. The maintainer must review this notice whenever the hosting model, storage behaviour, logging, analytics, or external integrations change.

This document describes the application's intended technical behaviour. It is not legal advice and does not replace an institution's own assessment of confidential examination processing.

