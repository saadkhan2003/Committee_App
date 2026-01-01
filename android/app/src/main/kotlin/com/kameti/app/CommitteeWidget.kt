package com.kameti.app

import android.appwidget.AppWidgetManager
import android.appwidget.AppWidgetProvider
import android.content.Context
import android.widget.RemoteViews
import es.antonborri.home_widget.HomeWidgetPlugin

class CommitteeWidget : AppWidgetProvider() {

    override fun onUpdate(
        context: Context,
        appWidgetManager: AppWidgetManager,
        appWidgetIds: IntArray
    ) {
        for (appWidgetId in appWidgetIds) {
            updateAppWidget(context, appWidgetManager, appWidgetId)
        }
    }

    override fun onEnabled(context: Context) {
        // Enter relevant functionality for when the first widget is created
    }

    override fun onDisabled(context: Context) {
        // Enter relevant functionality for when the last widget is disabled
    }

    companion object {
        internal fun updateAppWidget(
            context: Context,
            appWidgetManager: AppWidgetManager,
            appWidgetId: Int
        ) {
            // Get data from shared preferences (saved by Flutter)
            val widgetData = HomeWidgetPlugin.getData(context)
            
            val memberName = widgetData.getString("next_payout_member", "No data") ?: "No data"
            val payoutDate = widgetData.getString("next_payout_date", "") ?: ""
            val amount = widgetData.getString("payout_amount", "") ?: ""
            val committeeName = widgetData.getString("committee_name", "Committee") ?: "Committee"
            
            val views = RemoteViews(context.packageName, R.layout.committee_widget)
            
            views.setTextViewText(R.id.widget_title, committeeName)
            views.setTextViewText(R.id.widget_member_name, memberName)
            views.setTextViewText(R.id.widget_payout_date, payoutDate)
            views.setTextViewText(R.id.widget_amount, amount)
            
            appWidgetManager.updateAppWidget(appWidgetId, views)
        }
    }
}
