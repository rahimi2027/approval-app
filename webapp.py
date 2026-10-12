                                f.read(),
                                file_name=f"BACKUP_{stem}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                type="primary",
                                key=key,
                                width="stretch",
                            )
                    except Exception as e:
                        st.warning(f"Unable to prepare {label}: {e}")
                else:
                    st.button(f"{label} (not available)", disabled=True, key=f"{key}_missing")

        st.divider()
        st.subheader("🚀 Make Software Live")
        st.warning(
            "Use this only after downloading the backups above. "
            "It permanently clears all operational/test data: requests, work orders, "
            "inspector bonuses, HR leave requests and audit history. User accounts and system settings are NOT deleted."
        )

        def _clear_live_launch_data():
            clear_all_requests_file()
            clear_all_inspector_bonus()
            _write_empty_excel(WORK_ORDERS_PATH, WORK_ORDER_COLUMNS)
            _write_empty_excel(HR_LEAVE_PATH, HR_LEAVE_COLUMNS)
            _write_empty_excel(STORE_DEDUCTION_PATH, STORE_DEDUCTION_COLUMNS)
            _write_empty_excel(EMPLOYEE_ITEMS_PATH, EMPLOYEE_ITEMS_COLUMNS)
            _write_empty_excel(ITEM_CHECKIN_PATH, ITEM_CHECKIN_COLUMNS)
            _invalidate_data_cache(
                "_work_orders_cache",
                "_work_order_cache",
                "_inspector_bonus_cache",
                "_records_cache",
                "_hr_leave_cache",
                "_store_deduction_cache",
                "_employee_items_cache",
                "_item_checkins_cache",
            )
            sync_saved_file_to_drive(WORK_ORDERS_PATH)
            sync_saved_file_to_drive(HR_LEAVE_PATH)
            sync_saved_file_to_drive(STORE_DEDUCTION_PATH)
            sync_saved_file_to_drive(EMPLOYEE_ITEMS_PATH)
            sync_saved_file_to_drive(ITEM_CHECKIN_PATH)
            clear_audit_log_file()

            for folder in (PDF_DIR, WORK_ORDER_PDF_DIR, INSPECTOR_BONUS_PDF_DIR, HR_LEAVE_PDF_DIR, STORE_DEDUCTION_PDF_DIR, UPLOAD_DIR):
                if os.path.isdir(folder):
                    for root, dirs, files in os.walk(folder, topdown=False):
                        for filename in files:
                            try:
                                os.remove(os.path.join(root, filename))
                            except OSError:
                                pass
                        for dirname in dirs:
                            try:
                                os.rmdir(os.path.join(root, dirname))
                            except OSError:
                                pass
            st.session_state["_live_data_reset"] = datetime.now().isoformat()

        if not st.session_state.get("confirm_live_launch", False):
            if st.button(
                "🚀 Make Software Live — Clear All Operational Data",
                type="primary",
                key="make_software_live_btn",
                width="stretch",
            ):
                st.session_state["confirm_live_launch"] = True
                st.rerun()
        else:
            st.error(
                "⚠️ FINAL CONFIRMATION: this will permanently remove all requests, "
                "work orders, inspector bonuses, HR leave requests, Store Deductions, audit history, generated PDFs and uploaded attachments. "
                "Users and system settings will remain."
            )
            live_c1, live_c2 = st.columns(2)
            with live_c1:
                if st.button(
                    "🚀 YES — Make Software Live",
                    type="primary",
                    key="confirm_make_software_live_btn",
                    width="stretch",
                ):
                    try:
                        _clear_live_launch_data()
                        st.session_state["confirm_live_launch"] = False
                        st.success("✅ Software is ready for live use. All operational/test data has been cleared.")
                        st.rerun()
                    except Exception as e:
                        st.session_state["confirm_clear_live_launch"] = False
                        st.error(f"❌ Live launch reset failed: {e}")
            with live_c2:
                if st.button("↩️ Cancel", key="cancel_make_software_live_btn", width="stretch"):
                    st.session_state["confirm_live_launch"] = False
                    st.rerun()

    with super_training_help_tab:
        if user_info.get("can_access_training_help", False):
            render_training_help(user_info)
        else:
            st.info("Training & Help Centre is not enabled for this account. Please enable the module if required.")

else:
    st.subheader("🔐 Access Restricted")
    st.error("❌ Your role does not have a defined portal. Please contact Super Admin.")
# ========================================================
# ✅ END OF ROLE-BASED PORTALS
# ============================================================
# ✅ END OF FILE — NOTHING AFTER THIS!
# ============================================================

