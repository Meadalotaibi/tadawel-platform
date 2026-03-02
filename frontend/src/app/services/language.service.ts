import { Injectable, signal } from '@angular/core';

export type Language = 'en' | 'ar';

export interface Translations {
    title: string;
    subtitle: string;
    uploadPage: {
        title: string;
        saherLabel: string;
        saherPlaceholder: string;
        mvaLabel: string;
        mvaPlaceholder: string;
        analyzeButton: string;
        uploadInstruction: string;
        fileValidation: string;
    };
    dateRange: {
        title: string;
        startDate: string;
        endDate: string;
        hint: string;
        validation: string;
    };
    statistics: {
        title: string;
        cancels: string;
        totalViolations: string;
        totalViolationsSelected: string;
        weeklyViolations: string;
        cancelsSelected: string;
        selectedRange: string;
        saherViolationsTable: string;
        hsseViolationsTable: string;
        noData: string;
        total: string;
        count: string;
        businessLine: string;
        group: string;
        exportButton: string;
        newAnalysis: string;
    };
    language: {
        english: string;
        arabic: string;
    };
}

@Injectable({
    providedIn: 'root'
})
export class LanguageService {
    currentLanguage = signal<Language>('en');

    private translations: Record<Language, Translations> = {
        en: {
            title: 'Tadawel Report',
            subtitle: 'Upload and analyze your Excel data',
            uploadPage: {
                title: 'Upload Excel Files',
                saherLabel: 'Saher Violation File',
                saherPlaceholder: 'Select Saher violation Excel file',
                mvaLabel: 'MVA File',
                mvaPlaceholder: 'Select MVA Excel file',
                analyzeButton: 'Analyze Data',
                uploadInstruction: 'Upload both Excel files to begin analysis',
                fileValidation: 'Please upload both Excel files (.xlsx, .xls)'
            },
            dateRange: {
                title: 'Select Date Range',
                startDate: 'Start Date',
                endDate: 'End Date',
                hint: 'Select start and end date',
                validation: 'Please select both start and end dates'
            },
            statistics: {
                title: 'Statistics',
                cancels: 'Cancels',
                totalViolations: 'Total Violations',
                totalViolationsSelected: 'Total Violations',
                weeklyViolations: 'Weekly Violations (Wed–Tue)',
                cancelsSelected: 'Cancels',
                selectedRange: 'Selected Range',
                saherViolationsTable: 'SAHER Violations by Business Line & Region',
                hsseViolationsTable: 'HSSE Violations by Group & Region',
                noData: 'No data for selected range',
                total: 'Total',
                count: 'Count',
                businessLine: 'Business Line Org Description',
                group: 'Group',
                exportButton: 'Export Tadawel Files',
                newAnalysis: 'New Analysis'
            },
            language: {
                english: 'English',
                arabic: 'العربية'
            }
        },
        ar: {
            title: 'تقرير تداول',
            subtitle: 'قم برفع وتحليل بيانات Excel الخاصة بك',
            uploadPage: {
                title: 'رفع ملفات Excel',
                saherLabel: 'ملف مخالفات ساهر',
                saherPlaceholder: 'اختر ملف Excel لمخالفات ساهر',
                mvaLabel: 'ملف MVA',
                mvaPlaceholder: 'اختر ملف MVA Excel',
                analyzeButton: 'تحليل البيانات',
                uploadInstruction: 'قم برفع ملفي Excel للبدء في التحليل',
                fileValidation: 'الرجاء رفع ملفي Excel (.xlsx, .xls)'
            },
            dateRange: {
                title: 'اختر نطاق التاريخ',
                startDate: 'تاريخ البداية',
                endDate: 'تاريخ النهاية',
                hint: 'اختر تاريخ البداية والنهاية',
                validation: 'الرجاء تحديد تاريخ البداية والنهاية'
            },
            statistics: {
                title: 'الإحصائيات',
                cancels: 'الملغاة',
                totalViolations: 'إجمالي المخالفات',
                totalViolationsSelected: 'إجمالي المخالفات',
                weeklyViolations: 'مخالفات الأسبوع (أربعاء–ثلاثاء)',
                cancelsSelected: 'الملغاة',
                selectedRange: 'النطاق المحدد',
                saherViolationsTable: 'مخالفات ساهر حسب خط الأعمال والمنطقة',
                hsseViolationsTable: 'مخالفات الصحة والسلامة حسب المجموعة والمنطقة',
                noData: 'لا توجد بيانات للنطاق المحدد',
                total: 'المجموع',
                count: 'العدد',
                businessLine: 'وصف خط الأعمال',
                group: 'المجموعة',
                exportButton: 'تصدير ملفات تداول',
                newAnalysis: 'تحليل جديد'
            },
            language: {
                english: 'English',
                arabic: 'العربية'
            }
        }
    };

    getTranslations(): Translations {
        return this.translations[this.currentLanguage()];
    }

    switchLanguage(lang: Language) {
        this.currentLanguage.set(lang);
        // Update HTML dir attribute for RTL support
        document.documentElement.dir = lang === 'ar' ? 'rtl' : 'ltr';
        document.documentElement.lang = lang;
    }

    toggleLanguage() {
        const newLang: Language = this.currentLanguage() === 'en' ? 'ar' : 'en';
        this.switchLanguage(newLang);
    }
}
