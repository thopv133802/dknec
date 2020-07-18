package com.thopham.projects.dknec.demo_v2

import com.google.gson.Gson
import com.google.gson.reflect.TypeToken
import java.math.BigInteger
import java.sql.Timestamp
import java.util.concurrent.ThreadLocalRandom
import javax.persistence.Column
import javax.persistence.Entity
import javax.persistence.Id
import kotlin.collections.ArrayList

data class Tag(
        val tag_name: String = "",
        val tag_value: Float = 0.0f,
        val created: String = ""
) {
    companion object {
        fun fromSQL(tag_name: String, sql: Any): Tag{
            val sqlArray = sql as Array<Any>
            return Tag(tag_name, sqlArray[0] as Float, (sqlArray[1] as Timestamp).toString())
        }
    }
}