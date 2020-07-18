package com.thopham.projects.dknec.demo_v2

import org.springframework.beans.factory.annotation.Value
import org.springframework.stereotype.Component
import reactor.core.publisher.Flux
import reactor.core.publisher.Mono
import java.sql.Timestamp
import java.text.SimpleDateFormat
import java.time.Duration
import java.util.*
import java.util.concurrent.ThreadLocalRandom
import javax.annotation.PostConstruct

@Component
class TagFallback(val mesClient: MesClient) {
    @Value("\${interval_insecond:10}")
    var interval: Long = 0L
//    @PostConstruct
    fun run() {
        Flux.interval(Duration.ofSeconds(interval))
                .map {
                    val lastTag: Tag? = mesClient.getLastTag()
                    println("Last tag: $lastTag")
                    if(lastTag == null) return@map 0
                    val nextMinute = nextMinute(lastTag.created)
                    var tags = mesClient.getFirstTags().distinctBy { tag -> tag.tag_name }
                    println(tags.drop(120).take(2))
                    tags = tags.map { tag ->
                        if(tag.tag_value.compareTo(0) == 0 || tag.tag_value.compareTo(1) == 0) {
                            tag.copy(created = nextMinute)
                        }
                        else
                            tag.copy(created = nextMinute, tag_value = tag.tag_value + ThreadLocalRandom.current().nextFloat() / 10 * tag.tag_value)
                    }
                    println(tags.drop(120).take(2))
                    var addedAmount = 0
//                    if(tags.isNotEmpty()) {
//                        addedAmount = mesClient.addTags(AddTagsParams(tags = tags)).size
//                    }
                    addedAmount
                }
                .onErrorContinue { error, _ ->
                    println("Error 1: ${error.message}")
                }
                .subscribe({ addedAmount ->
                    println("Added amount: $addedAmount")
                }, {err ->
                    println("Error 2: ${err.message}")
                })
    }

    private fun nextMinute(example: String): String {
        val dateFormat = SimpleDateFormat("yyyy-MM-dd HH:mm:ss")
        val date = dateFormat.parse(example)
        val calendar = Calendar.getInstance()
        calendar.time = date
        calendar.add(Calendar.MINUTE, 1)
        val nextMinuteDate = Timestamp(calendar.timeInMillis)
        return nextMinuteDate.toString()
    }
}