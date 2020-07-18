package com.thopham.projects.dknec.demo_v2

import org.springframework.beans.factory.annotation.Value
import org.springframework.stereotype.Component
import reactor.core.publisher.Flux
import java.time.Duration
import javax.annotation.PostConstruct

@Component
class TagScheduler(val mesClient: MesClient, val service: TagService){
    @Value("\${interval_insecond:10}")
    var interval: Long = 0L
    @PostConstruct
    fun schedule() {
        Flux.interval(Duration.ofSeconds(interval))
                .map {
                    val lastTag: Tag? = mesClient.getLastTag()
                    println("Last tag: $lastTag")
                    val tags = if(lastTag != null)
                        service.findAllByCreatedGreaterThanOrderByCreatedAsc(lastTag.created)
                    else
                        service.findAllOrderByCreatedAsc()
                    var addedAmount = 0
                    if(tags.isNotEmpty()) {
                        addedAmount = mesClient.addTags(AddTagsParams(tags = tags)).size
                    }
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
}