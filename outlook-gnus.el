;;; outlook-gnus.el --- integration with `gnus' package  -*- lexical-binding: t; -*-

;; Copyright (C) 2018  Andrew Savonichev

;; Author: Andrew Savonichev
;; URL: https://github.com/asavonic/outlook.el
;; Version: 0.1
;; Keywords: mail

;; Package-Requires: ((emacs "24.4"))

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:

;; This is an integration package to use outlook.el with `gnus' MUA
;; package.
;;
;; Use `outlook-gnus-html-message-finalize' before sending a reply in
;; order to format it. Run `outlook-gnus-html-message-preview' to
;; preview the final message before sending it.

;;; Code:
(require 'outlook)
(require 'gnus)
(require 'gnus-art)
(require 'message)
(require 'ietf-drums)
(require 'subr)
(require 'subr-x)

(defun outlook-gnus-message-finalize ()
  (interactive)
  (let ((message   (outlook-gnus-parent-message))
        (html-body (outlook-gnus-parent-html-body))
        (txt-body  (outlook-gnus-parent-txt-body)))

    (cond
     (html-body
      (outlook-gnus-html-message-finalize message html-body))
     (txt-body
      (outlook-gnus-txt-message-finalize message txt-body))
     (t (error "Cannot find parent message body.")))))

(defun outlook-gnus-html-message-finalize (message html)
  (message-goto-body)
  (outlook-html-insert-reply html message (point) (point-max))
  (delete-region (point) (point-max))
  (insert "<#part type=\"text/html\">\n")
  (outlook-html-dom-print html))

(defun outlook-gnus-html-message-preview ()
  (interactive)
  (save-excursion
    (message-goto-body)

    (when (outlook-gnus-parent-html-body)
      (unless (search-forward "<#part type=\"text/html\">" nil t)
        (error "Did you call outlook-gnus-html-message-finalize before?")))

    (let ((temp-file (make-temp-file "emacs-email")))
      (write-region (point) (point-max) temp-file)
      (browse-url (concat "file://" temp-file)))))

(defun outlook-gnus-parent-html-body ()
  (let ((html-string (gnus-message-get-html-body)))
    (when html-string
      (with-temp-buffer
        (insert html-string)
        (outlook-html-read (point-min) (point-max))))))

(defun gnus-message-parent-message-id ()
  ;; TODO: fetch parent message using message-id
  (message-fetch-field "Message-ID" t))

(defun gnus-message-fetch-field (header &optional not-all)
  (with-current-buffer (or gnus-original-article-buffer
                           gnus-article-current-summary
                           (current-buffer))
    (message-fetch-field header not-all)))

(defun gnus-message-narrow-to-body ()
  (progn
    (message-goto-body)
    (narrow-to-region (point) (point-max))))

(defun gnus-message-get-html-body ()
  (with-current-buffer (or gnus-original-article-buffer
                           gnus-article-current-summary
                           (current-buffer))
    (progn
      (message-goto-body)
      (buffer-substring (point) (point-max)))))

(defun gnus-message-get-body ()
  (with-current-buffer (or gnus-original-article-buffer
                           gnus-article-current-summary
                           (current-buffer))
    (with-current-buffer (gnus-copy-article-buffer)
      (gnus-msg-treat-broken-reply-to)
      (save-restriction
        (gnus-message-narrow-to-body)
        (buffer-string)))))

(defun outlook-gnus-parent-message ()
  (outlook-message
   (outlook-gnus-format-sender
    (gnus-message-fetch-field "From" t))

   (outlook-gnus-format-recipients-list
    (gnus-message-fetch-field "To" t))

   (outlook-gnus-format-recipients-list
    (gnus-message-fetch-field "Cc" t))

   (outlook-format-date-string
    (ietf-drums-parse-date (gnus-message-fetch-field "Date" t)))

   (gnus-message-fetch-field "Subject" t)))

(defun outlook-gnus-parse-email (email)
  (let ((e (ietf-drums-parse-address email)))
    (cons (cdr e) (car e))))

(defun outlook-gnus-format-sender (contact)
  (when contact
    (let ((email (outlook-gnus-parse-email contact)))
      (outlook-format-sender (car email) (cdr email)))))

(defun outlook-gnus-format-recipients-list (contacts)
  (when contacts
    (let ((emails (mapcar #'outlook-gnus-parse-email
                          (split-string contacts ", "))))
     (string-join
      (mapcar (lambda (name-email)
                (outlook-format-recipient (car name-email) (cdr name-email)))
              emails)
      "; "))))

(defun outlook-gnus-parent-txt-body ()
  (gnus-message-get-body))

(defun outlook-gnus-txt-message-finalize (message txt)
  (save-excursion
    (goto-char (point-max))
    (insert "\n\n")
    (outlook-txt-insert-quote-header message)
    (insert "\n")
    (insert txt)))

(provide 'outlook-gnus)
;;; outlook-gnus.el ends here

